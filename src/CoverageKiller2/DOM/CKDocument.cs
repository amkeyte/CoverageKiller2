using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a wrapper for a Word document, managing its lifecycle and interactions
    /// with the Word application, including event handling for opening and closing.
    /// </summary>
    public class CKDocument : IDOMObject
    {
        private readonly string _fullPath;
        private bool documentOpened = false;

        /// <summary>
        /// Gets the associated Word document.
        /// </summary>
        public Word.Document COMDocument { get; private set; }

        /// <summary>
        /// Gets the Word application instance that is managing this document.
        /// </summary>
        public Word.Application Application => COMDocument.Application;

        /// <summary>
        /// Gets the content of the Word document as a <see cref="Word.Range"/>.
        /// </summary>
        public Word.Range Content => COMDocument.Content;

        /// <summary>
        /// Gets the full file path of the Word document.
        /// </summary>
        public string FullPath => _fullPath;




        // Using Create why? Probably just a conventional to prevent uninttended changes from other code.
        // if this is the case, probably better to switch to copying around the single refernce and then 
        // copying the range content if it makes sense in some case. Hell, this could even be for some old
        // shit im not even doing anymore.
        public CKTables Tables => throw new NotImplementedException();
        public CKSections Sections => new CKSections(Range());

        //circular; find a way to fix.
        public CKDocument Document => this;


        public IDOMObject Parent => throw new NotSupportedException("Call Application on a CKDocument object.");

        public bool IsDirty => throw new NotImplementedException();

        /// <summary>
        /// Gets a value indicating whether this CKDocument no longer has a valid COMDocument reference.
        /// This becomes true if the document is closed or the COM object has been released.
        /// </summary>
        public bool IsOrphan
        {
            get
            {
                try
                {
                    // Accessing COMDocument.Application should throw if the COM object is no longer valid.
                    // Alternatively, accessing COMDocument.FullName is often sufficient.
                    _ = COMDocument.FullName;
                    return false;
                }
                catch (COMException)
                {
                    return true;
                }
                catch (Exception)
                {
                    return true;
                }
            }
        }
        public CKRange Range() => new CKRange(COMDocument.Range(), this);

        public CKRange Range(Word.Range range, IDOMObject parent = null)
        {
            return new CKRange(COMDocument.Range(), parent);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="CKDocument"/> class, 
        /// opening the Word document located at the specified <paramref name="fullPath"/>.
        /// </summary>
        /// <param name="fullPath">The full path to the Word document.</param>
        public CKDocument(string fullPath)
        {
            _fullPath = fullPath;
            COMDocument = Open(fullPath);
            Log.Debug("Registering BeforeClose event for document {DocName}", COMDocument.FullName);
            COMDocument.Application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKDocument"/> class using an existing
        /// Word document instance.
        /// </summary>
        /// <param name="wordDoc">The existing Word document instance.</param>
        public CKDocument(Word.Document wordDoc)
        {
            COMDocument = wordDoc;
            _fullPath = COMDocument.FullName;
            documentOpened = true;

            Log.Debug("Registering BeforeClose event for document {DocName}", COMDocument.FullName);
            COMDocument.Application.DocumentBeforeClose += OnDocumentBeforeClose;
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
        public void Activate() => COMDocument.Activate();

        /// <summary>
        /// Closes the Word document. Optionally saves changes.
        /// </summary>
        /// <param name="saveChanges">If true, saves the changes before closing.</param>
        public virtual void Close(bool saveChanges = false)
        {
            COMDocument.Close(saveChanges);
        }

        /// <summary>
        /// Event handler that is triggered before the document is closed.
        /// </summary>
        /// <param name="wordDoc">The document being closed.</param>
        /// <param name="Cancel">Cancel the closing operation if set to true.</param>
        private void OnDocumentBeforeClose(Word.Document wordDoc, ref bool Cancel)
        {
            if (!IsThisDocument(wordDoc)) return;

            PurgeGlobalRefernces();

            Log.Information("Closed document {DocName}", wordDoc.FullName);
            Log.Debug("Unregistering BeforeClosed event for {DocName}", wordDoc.FullName);

            COMDocument.Application.DocumentBeforeClose -= OnDocumentBeforeClose;
        }

        private void PurgeGlobalRefernces()
        {
            CKTableGrid.PurgeInstances(this);
        }

        public Tracer Tracer = new Tracer(typeof(CKDocument));
        /// <summary>
        /// Deletes a specified section from the Word document.
        /// </summary>
        /// <param name="sectionIndex">The index of the section to delete (1-based).</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the section index is out of range.</exception>

        public void DeleteSection(int sectionIndex)
        {
            Tracer.Log("Deleting Section", new DataPoints()
                .Add(nameof(sectionIndex), sectionIndex));

            if (sectionIndex < 1 || sectionIndex > COMDocument.Sections.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Section index is out of range.");
            }

            // Get the section to delete
            Word.Section sectionToDelete = COMDocument.Sections[sectionIndex];

            // Delete the section and hopefully he ssection break ahead of it.
            Word.Range extendedRange = COMDocument.Range(sectionToDelete.Range.Start - 1, sectionToDelete.Range.End);

            //sectionToDelete.Range.Delete();
            extendedRange.Delete();
        }

        // Get primary footer range
        public Word.Range GetFooterRange()
        {
            return COMDocument.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        }

        // Get primary header range
        public Word.Range GetHeaderRange()
        {
            return COMDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        }

        // Copy footer from this document to another
        public void CopyFooterTo(CKDocument targetDocument)
        {
            if (targetDocument == null || targetDocument.COMDocument == null)
                throw new ArgumentNullException(nameof(targetDocument));

            Word.Range sourceFooter = GetFooterRange();
            Word.Range targetFooter = targetDocument.GetFooterRange();

            targetFooter.FormattedText = sourceFooter.FormattedText;
        }

        // Copy header from this document to another
        public void CopyHeaderTo(CKDocument targetDocument)
        {
            if (targetDocument == null || targetDocument.COMDocument == null)
                throw new ArgumentNullException(nameof(targetDocument));

            Word.Range sourceHeader = GetHeaderRange();
            Word.Range targetHeader = targetDocument.GetHeaderRange();

            targetHeader.FormattedText = sourceHeader.FormattedText;
        }

        // Copy both header and footer
        public void CopyHeaderAndFooterTo(CKDocument targetDocument)
        {
            CopyHeaderTo(targetDocument);
            CopyFooterTo(targetDocument);
        }

        internal CKRange Range(int start, int end, IDOMObject parent = null)
        {
            return Range(COMDocument.Range(start, end), parent ?? this);
        }
    }

}
