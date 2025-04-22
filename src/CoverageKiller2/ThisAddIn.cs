using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using Serilog;
using System;

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
        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                //ThisAddin is hijacking CKOffice in a new process it seems. remove comment to return

                if (!CKOffice_Word.IsTest) //kinda hacky. ThisAddin is hijacking CKOffice in a new process it seems.
                {

                    var OfficeWord = CKOffice_Word.Instance;
                    OfficeWord.Start();
                    OfficeWord.TryPutAddin(this);
                    LogExpertLoader.StartLogExpert(LoggingLoader.LogFile, true);
                    //string logFile = LogTailLoader.GetLogFile();
                    //LoggingLoader.Configure(logFile, Serilog.Events.LogEventLevel.Debug);

                    ////debugging the big hangup.
                    ////LogTailLoader.StartBareTail(logFile);

                    //Log.Debug("Logging started: Level {logEventLevel}", LoggingLoader.Level);
                    Log.Information("ThisAddIn started.");
                }
            }
            catch (Exception ex)
            {

                throw ex;

            }
        }

        /// <summary>
        /// Cleans up logging and BareTail when the add-in is shut down.
        /// </summary>
        /// <param name="sender">The event source.</param>
        /// <param name="e">Event arguments.</param>
        private async void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            ///All this stuff can go in CKOffice.ShutDown
            CKOffice_Word.Instance.ShutDown();

            //try
            //{
            //    Log.Information("ThisAddIn shutting down.");
            //    LoggingLoader.Cleanup();
            //    LogTailLoader.Cleanup();
            //}
            //catch (Exception ex)
            //{

            //    throw ex;

            //}
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
