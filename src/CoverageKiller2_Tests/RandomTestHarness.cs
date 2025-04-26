using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Test
{
    /// <summary>
    /// Centralized test harness for initializing CKOffice_Word and managing test documents.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0005
    /// </remarks>
    public static class RandomTestHarness
    {
        public static bool RandomTestHarnessIsRunning { get; private set; } = false;

        /// <summary>
        /// If true, temporary files will be preserved after test execution.
        /// </summary>
        public static bool PreserveTempFilesAfterTest { get; set; } = false;

        static RandomTestHarness()
        {
            InitializeOnceAtStartup();
        }

        private static void InitializeOnceAtStartup()
        {
            RandomTestHarnessIsRunning = true;
            CKOffice_Word.Instance.Start();
            LogExpertLoader.StartLogExpert(LoggingLoader.LogFile, true);
        }

        private static CKApplication _sharedApp;
        public static string TestFile1 = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (Noise Floor)_Test3.docx";
        public static string TestFile2 = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (CC) Short.docx";

        /// <summary>
        /// Returns a temp document created as a renamed copy of the given source,
        /// opened in a fresh or shared application instance with suppressed alerts.
        /// </summary>
        /// <param name="sourcePath">The original document to copy.</param>
        /// <param name="visible">Whether the document should be visible.</param>
        /// <param name="cleanApp">If true, uses a new application instance; otherwise uses the shared one.</param>
        /// <param name="filename">Optional filename override.</param>
        /// <returns>A test-safe <see cref="CKDocument"/> instance.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0005
        /// </remarks>
        public static CKDocument GetTempDocumentFrom(string sourcePath, bool visible = false, bool cleanApp = false, string filename = null)
        {
            filename = filename ?? Path.GetRandomFileName() + Path.GetExtension(sourcePath);
            var tempPath = Path.Combine(Path.GetTempPath(), filename);

            if (File.Exists(tempPath)) File.Delete(tempPath);
            File.Copy(sourcePath, tempPath);

            bool forceVisible = PreserveTempFilesAfterTest || visible;

            CKApplication newApp = default;

            if (_sharedApp == null)
            {
                _sharedApp = CKOffice_Word.Instance.TryGetNewApp(out newApp, forceVisible) >= 0
                    ? newApp
                    : throw new Exception("Failed to get clean Word app");
            }

            CKApplication app = cleanApp
                ? (CKOffice_Word.Instance.TryGetNewApp(out newApp, forceVisible) >= 0 ? newApp : throw new Exception("Failed to get clean Word app"))
                : _sharedApp ?? throw new NullReferenceException("_sharedApp is null");

            app.WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.WordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var doc = app.GetDocument(tempPath, visible: forceVisible, createIfNotFound: true);

            if (PreserveTempFilesAfterTest)
            {
                doc.KeepAlive = true;
            }

            return doc;
        }

        /// <summary>
        /// Shuts down the harness and all associated application instances.
        /// </summary>
        public static void Shutdown()
        {
            LogExpertLoader.Cleanup();
            CKOffice_Word.Instance.ShutDown();
            RandomTestHarnessIsRunning = false;
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
        /// Version: CK2.00.01.0005
        /// </remarks>
        public static void CleanUp(CKDocument doc, bool force = false)
        {
            if (doc == null) return;

            if (PreserveTempFilesAfterTest)
            {
                Log.Information($"Preserving temp document: {doc.FileName}");
                return;
            }

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
