using Serilog;
using System;
using System.Threading;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKDocument
    {
        private readonly string _fullPath;

        public Word.Document WordDoc { get; private set; }
        public Word.Application WordApp => WordDoc.Application;
        public Word.Range Content => WordDoc.Content;
        /// <summary>
        /// Use Path class naming conventions because Office DOM is all over the place.
        /// </summary>
        public string FullPath => _fullPath;

        public Word.Tables Tables => WordDoc.Tables;

        public CKDocument(string fullPath)
        {
            _fullPath = fullPath;
            WordDoc = Open(fullPath);
            Log.Debug("Registering BeforeClose event for document {DocName}", WordDoc.FullName);
            WordDoc.Application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        private Word.Document Open(string fullPath)
        {
            Log.Information("Accessing Word Document {fullPath}", fullPath);
            // Subscribe to the DocumentOpen event
            Globals.ThisAddIn.Application.DocumentOpen += OnDocumentOpen;
            Word.Document openedDoc = default;

            try
            {
                // Attempt to open the document
                openedDoc = Globals.ThisAddIn.Application.Documents.Open(
                     FileName: fullPath,
                     AddToRecentFiles: false,
                     ReadOnly: true,
                     Visible: false);

                // Wait for the document to finish loading
                int sleepTime = 100;
                int totalSleepTime = 0;
                // Wait until the documentOpened flag is set
                while (!documentOpened)
                {
                    Thread.Sleep(sleepTime); // Sleep for a short duration to avoid freezing
                    Forms.Application.DoEvents(); // Allow UI to remain responsive
                    totalSleepTime += sleepTime;
                }
                Log.Debug("Time to load {totalSleepTime} ms", totalSleepTime);

                // Continue processing here
                Log.Information("Document access success.");
            }
            catch (Exception ex)
            {
                Log.Error("Error opening document: {Message}", ex.Message);
                Forms.MessageBox.Show($"Error opening document: {ex.Message}");
                throw ex;
            }
            finally
            {
                // Unsubscribe from the event to avoid memory leaks
                Globals.ThisAddIn.Application.DocumentOpen -= OnDocumentOpen;
            }
            return openedDoc;

        }

        private bool documentOpened = false;
        private bool IsThisDocument(Word.Document wDoc) => wDoc.FullName == FullPath;

        private void OnDocumentOpen(Word.Document doc)
        {
            if (!IsThisDocument(doc)) return;
            documentOpened = true; // Set the flag when the document is opened
        }


        public CKDocument(Word.Document wordDoc)
        {
            WordDoc = wordDoc;
            _fullPath = WordDoc.FullName;

            documentOpened = true;
            Log.Debug("Registering BeforeClose event for document {DocName}", WordDoc.FullName);
            WordDoc.Application.DocumentBeforeClose += OnDocumentBeforeClose;

        }

        public void Activate() => WordDoc.Activate();

        public virtual void Close(bool saveChanges = false)
        {
            WordDoc.Close(saveChanges);
        }

        private void OnDocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            if (!IsThisDocument(Doc)) return;

            Log.Information("Closed document{DocName}", Doc.FullName);

            Log.Debug("Unregistering BeforeClosed event for {DocName}", Doc.FullName);
            WordDoc.Application.DocumentBeforeClose -= OnDocumentBeforeClose;
        }

    }
}
