using CoverageKiller2.Logging;
using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2.Tests
{
    /// <summary>
    /// Helper class for integration tests that require a live Word instance.
    /// It loads a document from disk and returns the Word.Document DOM.
    /// Remember to dispose of the instance to clean up the Word process.
    /// </summary>
    public class LiveWordDocument : IDisposable
    {
        public const string Default = @"C:\Users\akeyte.PCM\source\repos\CoverageKiller2\src\CoverageKiller2_Tests\TestFiles\SEA Garage (Noise Floor)_20250313_152027.docx";



        private Word.Application _wordApp;

        public LiveWordDocument()
        {
            // Initialize the Word Application.
            _wordApp = new Word.Application
            {
                // Hide the UI during tests.
                Visible = false
            };
        }

        /// <summary>
        /// Loads a document from the specified file path.
        /// Opens the document in read-only mode.
        /// </summary>
        /// <param name="documentPath">Full file path to the document.</param>
        /// <returns>A live Word.Document object for testing.</returns>
        public Word.Document LoadFromFile(string documentPath)
        {
            LH.Ping(GetType());
            if (string.IsNullOrEmpty(documentPath))
            {
                throw new ArgumentNullException(nameof(documentPath));
            }

            // Open the document in read-only mode.
            Word.Document doc = _wordApp.Documents.Open(documentPath,
                                                          ReadOnly: true,
                                                          Visible: false);
            LH.Pong(GetType());
            return doc;
        }

        /// <summary>
        /// Closes the specified document and releases its COM object.
        /// </summary>
        /// <param name="document">The Word.Document to close.</param>
        public void Close(Word.Document document)
        {
            LH.Ping(GetType());

            if (document != null)
            {
                document.Close(false);
                Marshal.ReleaseComObject(document);
            }
            LH.Pong(GetType());

        }

        /// <summary>
        /// Disposes the Word Application instance.
        /// </summary>
        public void Dispose()
        {
            LH.Ping(GetType());

            if (_wordApp != null)
            {
                try
                {
                    _wordApp.Quit();
                }
                catch { /* Ignore any exceptions on quitting */ }
                finally
                {
                    Marshal.ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            }
            LH.Pong(GetType());

        }

        /// <summary>
        /// Loads the specified document, executes the given test action, and ensures the document is closed.
        /// </summary>
        /// <param name="documentPath">The full path to the test document.</param>
        /// <param name="testAction">The action to perform using the loaded Word.Document.</param>
        public static void WithTestDocument(string documentPath, Action<Word.Document> testAction)
        {
            LH.Ping(typeof(LiveWordDocument));

            using (var loader = new LiveWordDocument())
            {
                if (!string.IsNullOrEmpty(documentPath)) documentPath = Default;

                Word.Document doc = loader.LoadFromFile(documentPath);
                try
                {
                    testAction(doc);
                }
                finally
                {
                    loader.Close(doc);
                }
            }
            LH.Pong(typeof(LiveWordDocument));

        }

    }
}
