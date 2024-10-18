using Serilog;
using System;
using System.Threading;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents a wrapper for a Word document, managing its lifecycle and interactions
    /// with the Word application, including event handling for opening and closing.
    /// </summary>
    public class CKDocument
    {
        private readonly string _fullPath;
        private bool documentOpened = false;

        /// <summary>
        /// Gets the associated Word document.
        /// </summary>
        public Word.Document WordDoc { get; private set; }

        /// <summary>
        /// Gets the Word application instance that is managing this document.
        /// </summary>
        public Word.Application WordApp => WordDoc.Application;

        /// <summary>
        /// Gets the content of the Word document as a <see cref="Word.Range"/>.
        /// </summary>
        public Word.Range Content => WordDoc.Content;

        /// <summary>
        /// Gets the full file path of the Word document.
        /// </summary>
        public string FullPath => _fullPath;

        /// <summary>
        /// Gets the collection of tables in the Word document.
        /// </summary>
        public CKTables Tables => new CKTables(WordDoc.Tables);



        public CKTables CKTables { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKDocument"/> class, 
        /// opening the Word document located at the specified <paramref name="fullPath"/>.
        /// </summary>
        /// <param name="fullPath">The full path to the Word document.</param>
        public CKDocument(string fullPath)
        {
            _fullPath = fullPath;
            WordDoc = Open(fullPath);
            CKTables = new CKTables(this);
            Log.Debug("Registering BeforeClose event for document {DocName}", WordDoc.FullName);
            WordDoc.Application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKDocument"/> class using an existing
        /// Word document instance.
        /// </summary>
        /// <param name="wordDoc">The existing Word document instance.</param>
        public CKDocument(Word.Document wordDoc)
        {
            WordDoc = wordDoc;
            _fullPath = WordDoc.FullName;
            documentOpened = true;
            CKTables = new CKTables(this);

            Log.Debug("Registering BeforeClose event for document {DocName}", WordDoc.FullName);
            WordDoc.Application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        /// <summary>
        /// Opens a Word document at the specified path and waits for it to fully load.
        /// </summary>
        /// <param name="fullPath">The full path to the Word document.</param>
        /// <returns>The opened <see cref="Word.Document"/>.</returns>
        private Word.Document Open(string fullPath)
        {
            Log.Information("Accessing Word Document {fullPath}", fullPath);
            Globals.ThisAddIn.Application.DocumentOpen += OnDocumentOpen;
            Word.Document openedDoc = null;

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
                    Thread.Sleep(sleepTime);
                    Forms.Application.DoEvents();
                    totalSleepTime += sleepTime;
                }

                Log.Debug("Time to load {totalSleepTime} ms", totalSleepTime);
                Log.Information("Document access success.");
            }
            catch (Exception ex)
            {
                Log.Error("Error opening document: {Message}", ex.Message);
                Forms.MessageBox.Show($"Error opening document: {ex.Message}");
                throw;
            }
            finally
            {
                Globals.ThisAddIn.Application.DocumentOpen -= OnDocumentOpen;
            }

            return openedDoc;
        }

        /// <summary>
        /// Determines whether the specified Word document matches this document.
        /// </summary>
        /// <param name="wDoc">The Word document to compare.</param>
        /// <returns>True if the documents match, otherwise false.</returns>
        private bool IsThisDocument(Word.Document wDoc) => wDoc.FullName == FullPath;

        /// <summary>
        /// Event handler that is triggered when a Word document is opened.
        /// </summary>
        /// <param name="doc">The opened Word document.</param>
        private void OnDocumentOpen(Word.Document doc)
        {
            if (IsThisDocument(doc))
            {
                documentOpened = true;
            }
        }

        /// <summary>
        /// Activates the Word document in the application.
        /// </summary>
        public void Activate() => WordDoc.Activate();

        /// <summary>
        /// Closes the Word document. Optionally saves changes.
        /// </summary>
        /// <param name="saveChanges">If true, saves the changes before closing.</param>
        public virtual void Close(bool saveChanges = false)
        {
            WordDoc.Close(saveChanges);
        }

        /// <summary>
        /// Event handler that is triggered before the document is closed.
        /// </summary>
        /// <param name="wordDoc">The document being closed.</param>
        /// <param name="Cancel">Cancel the closing operation if set to true.</param>
        private void OnDocumentBeforeClose(Word.Document wordDoc, ref bool Cancel)
        {
            if (!IsThisDocument(wordDoc)) return;

            Log.Information("Closed document {DocName}", wordDoc.FullName);
            Log.Debug("Unregistering BeforeClosed event for {DocName}", wordDoc.FullName);

            WordDoc.Application.DocumentBeforeClose -= OnDocumentBeforeClose;
        }


        /// <summary>
        /// Deletes a specified section from the Word document.
        /// </summary>
        /// <param name="sectionIndex">The index of the section to delete (1-based).</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the section index is out of range.</exception>
        public void DeleteSection(int sectionIndex)
        {
            if (sectionIndex < 1 || sectionIndex > WordDoc.Sections.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Section index is out of range.");
            }

            // Get the section to delete
            Word.Section sectionToDelete = WordDoc.Sections[sectionIndex];

            // Delete the section
            sectionToDelete.Range.Delete();
        }
    }
}
