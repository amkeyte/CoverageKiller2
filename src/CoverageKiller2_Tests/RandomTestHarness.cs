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

        static RandomTestHarness()
        {
            CKOffice_Word.Instance.Start();
            if (CKOffice_Word.Instance.TryGetNewApp(out var app, visible: false) < 0)
                throw new Exception("Failed to acquire shared test CKApplication");
            _sharedApp = app;
        }

        /// <summary>
        /// Gets a document from the shared application.
        /// </summary>
        public static CKDocument GetDocument(string fullPath)
            => _sharedApp.GetDocument(fullPath, visible: false);

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
        public static CKDocument GetTempDocument()
            => _sharedApp.GetTempDocument();

        /// <summary>
        /// Returns a temp document created as a copy of a given source.
        /// </summary>
        public static CKDocument GetTempDocumentFrom(string sourcePath, bool cleanApp = false)
        {
            var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + Path.GetExtension(sourcePath));
            File.Copy(sourcePath, tempPath);
            return cleanApp
                ? GetIsolatedDocument(tempPath)
                : GetDocument(tempPath);
        }

        /// <summary>
        /// Shuts down the harness and all associated application instances.
        /// </summary>
        public static void Shutdown()
        {
            CKOffice_Word.Instance.ShutDown();
        }

        /// <summary>
        /// Gets the shared application (only use if necessary).
        /// </summary>
        public static CKApplication Application => _sharedApp;
    }
}
