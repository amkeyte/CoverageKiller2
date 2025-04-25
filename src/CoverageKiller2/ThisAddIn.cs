using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Linq;
namespace CoverageKiller2
{

    //    stupid addin won't activate now

    //options -> addins -> manage com -> checkmark won't stay checked.


    //ChatGPT said:
    //That's a classic and frustrating VSTO issue — let’s walk through what usually causes that "checkmark won’t stay checked" behavior, and how to fix it.

    //🔧 Step-by-Step Fix Checklist
    //✅ 1. Check for Startup Errors
    //Word disables add-ins that throw during Startup.

    //Even logging failures, async code, or file locks can trigger this.

    //🛠️ Fix:

    //Open Windows Registry Editor (regedit)

    //Go to:

    //php-template
    //Copy
    //Edit
    //HKEY_CURRENT_USER\Software\Microsoft\Office\<version>\Word\Resiliency\DisabledItems
    //Delete the binary value related to your add-in (or all if you're not sure)

    //Then try re-checking the box in Word → COM Add-ins.



    //##################################################################
    //keep in csproj file if word quits responding to start.
    //<PropertyGroup>
    //        <StartAction>Program</StartAction>
    //        <StartProgram>C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE</StartProgram>
    //</PropertyGroup> 
    //####################################################################


    /// <summary>
    /// The main class for the VSTO add-in. Handles initialization, shutdown,
    /// and custom ribbon functionality for the Word add-in.
    /// </summary>
    public partial class ThisAddIn
    {

        private bool IsRunningUnderTest()
        {
            return AppDomain.CurrentDomain.GetAssemblies()
                .Any(a =>
                    a.FullName.StartsWith("Microsoft.VisualStudio.TestPlatform", StringComparison.OrdinalIgnoreCase) ||
                    a.FullName.IndexOf("testhost", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    a.FullName.IndexOf("Microsoft.VisualStudio.QualityTools", StringComparison.OrdinalIgnoreCase) >= 0
                );
        }

        /// <summary>
        /// Initializes logging and BareTail when the add-in starts.
        /// </summary>
        /// <param name="sender">The event source.</param>
        /// <param name="e">Event arguments.</param>
        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //skip out if tests are running.
            if (IsRunningUnderTest())
            {
                Log.Information("Skipping add-in startup because test environment is detected.");
                return;
            }

            var OfficeWord = CKOffice_Word.Instance;
            if (OfficeWord.Start() + OfficeWord.TryPutAddin(this) == 0)
            {
                RegisterCrashHandlers();
                LogExpertLoader.StartLogExpert(LoggingLoader.LogFile, true);
                Log.Information("ThisAddIn started.");
            }
            else
            {
                Log.Warning("ThisAddin was refused control of CKOffice_Word");
            }


        }
        private void RegisterCrashHandlers()
        {
            // Unhandled Word errors
            this.Application.DocumentChange += () =>
            {
                try
                {
                    // Your safety check or doc validation logic here
                }
                catch (Exception ex)
                {
                    CKOffice_Word.Instance.Crash(ex, typeof(ThisAddIn), nameof(Application.DocumentChange));
                }
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                CKOffice_Word.Instance.Crash(ex, typeof(ThisAddIn), "AppDomain.UnhandledException");
            };

            Application.DocumentOpen += (doc) =>
            {
                try
                {
                    // Optional additional hook
                }
                catch (Exception ex)
                {
                    CKOffice_Word.Instance.Crash(ex, typeof(ThisAddIn), nameof(Application.DocumentOpen));
                }
            };
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
