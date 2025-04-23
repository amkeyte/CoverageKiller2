
using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using System;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Test
{
    /// <summary>
    /// Centralized test harness for initializing CKOffice_Word and managing test documents.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public static class RandomTestHarness
    {
        private static CKApplication _sharedApp;
        public static string TestFile1 = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (Noise Floor)_Test3.docx";
        public static string TestFile2 = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (CC) Short.docx";

        static RandomTestHarness()
        {
            CKOffice_Word.Instance.Start();
            LogExpertLoader.StartLogExpert(LoggingLoader.LogFile, restartIfOpen: true);

            if (CKOffice_Word.Instance.TryGetNewApp(out var app, visible: true) < 0)
                throw new Exception("Failed to acquire shared test CKApplication");
            _sharedApp = app;
        }

        /// <summary>
        /// Gets a document from the shared application.
        /// </summary>
        public static CKDocument GetDocument(string fullPath)
            => _sharedApp.GetDocument(fullPath, visible: true);

        /// <summary>
        /// Gets a document from a specific application index.
        /// </summary>
        public static CKDocument GetDocumentFromApp(int appIndex, string fullPath)
        {
            var app = CKOffice_Word.Instance.Applications.ElementAtOrDefault(appIndex);
            if (app == null) throw new IndexOutOfRangeException($"No CKApplication at index {appIndex}");
            return app.GetDocument(fullPath, visible: false);
        }

        /// <summary>
        /// Gets a document using a fresh Word application instance.
        /// </summary>
        public static CKDocument GetIsolatedDocument(string fullPath)
        {
            if (CKOffice_Word.Instance.TryGetNewApp(out var newApp, visible: false) < 0)
                throw new Exception("Could not acquire isolated CKApplication");
            return newApp.GetDocument(fullPath, visible: false);
        }

        /// <summary>
        /// Creates and returns a new temp document from the shared application.
        /// </summary>
        //public static CKDocument GetTempDocument()
        //    => _sharedApp.GetTempDocument();

        /// <summary>
        /// Returns a temp document created as a renamed copy of the given source,
        /// opened in a fresh or shared application instance with suppressed alerts.
        /// </summary>
        /// <param name="sourcePath">The original document to copy.</param>
        /// <param name="cleanApp">If true, uses a new application instance; otherwise uses the shared one.</param>
        /// <returns>A test-safe <see cref="CKDocument"/> instance.</returns>
        /// <remarks>
        /// Version: CK2.00.00.0007
        /// </remarks>
        public static CKDocument GetTempDocumentFrom(string sourcePath, bool visible = false, bool cleanApp = false)
        {
            var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + Path.GetExtension(sourcePath));
            File.Copy(sourcePath, tempPath);

            // Choose your app
            CKApplication app = cleanApp
                ? CKOffice_Word.Instance.TryGetNewApp(out var newApp, visible) >= 0 ? newApp : throw new Exception("Failed to get clean Word app")
                : _sharedApp;

            // Apply safety settings (silence alerts and disable macro warnings)
            app.WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.WordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            return app.GetDocument(tempPath, visible: false);
        }


        /// <summary>
        /// Shuts down the harness and all associated application instances.
        /// </summary>
        public static void Shutdown()
        {
            LogExpertLoader.Cleanup();
            CKOffice_Word.Instance.ShutDown();
        }

        /// <summary>
        /// Gets the shared application (only use if necessary).
        /// </summary>
        public static CKApplication Application => _sharedApp;
        /// <summary>
        /// Cleans up the specified document using its owning application.
        /// </summary>
        /// <param name="doc">The document to clean up.</param>
        /// <param name="force">If true, disposes even orphaned documents.</param>
        /// <remarks>
        /// Version: CK2.00.00.0003
        /// </remarks>
        public static void CleanUp(CKDocument doc, bool force = false)
        {
            if (doc == null) return;
            try
            {
                doc.Application.CloseDocument(doc, force);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CleanUp] Error cleaning up document {doc.FullPath}: {ex.Message}");
            }
        }
    }
}
