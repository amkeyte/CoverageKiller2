using Serilog;
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a Word application instance responsible for creating, managing,
    /// and closing documents in a controlled environment.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0002
    /// </remarks>
    public class CKApplication : IDisposable
    {
        private readonly Word.Application _wordApp;
        private readonly List<CKDocument> _documents = new List<CKDocument>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CKApplication"/> class with an existing Word.Application.
        /// </summary>
        /// <param name="wordApp">The Word application instance to wrap.</param>
        public CKApplication(Word.Application wordApp)
        {
            _wordApp = wordApp ?? throw new ArgumentNullException(nameof(wordApp));
            Log.Information("CKApplication created.");
        }

        /// <summary>
        /// The raw Word application instance.
        /// </summary>
        public Word.Application WordApp => _wordApp;//<GPT no more public COM objects.

        /// <summary>
        /// Indicates whether this application is the VSTO ThisAddIn instance.
        /// </summary>
        public bool IsAddIn => ReferenceEquals(_wordApp, Globals.ThisAddIn?.Application);

        /// <summary>
        /// Gets the documents currently open in this application.
        /// </summary>
        public IReadOnlyList<CKDocument> Documents => _documents.AsReadOnly();

        /// <summary>
        /// Opens a document and wraps it in a <see cref="CKDocument"/>.
        /// </summary>
        /// <param name="fullPath">The full file path to the document.</param>
        /// <param name="visible">Whether the document should be visible when opened.</param>
        /// <returns>A new <see cref="CKDocument"/> wrapping the opened Word document.</returns>
        public CKDocument GetDocument(string fullPath, bool visible)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
                throw new ArgumentException("Invalid file path.", nameof(fullPath));

            try
            {
                Log.Information("Opening document from path: {Path}", fullPath);
                var doc = _wordApp.Documents.Open(
                    FileName: fullPath,
                    ReadOnly: false,
                    Visible: visible
                );

                var ckDoc = new CKDocument(doc, this);
                _documents.Add(ckDoc);
                Log.Information("Document opened and tracked: {Path}", fullPath);
                return ckDoc;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to open document {Path}: {Message}", fullPath, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Closes and removes the specified document from the application.
        /// </summary>
        /// <param name="doc">The document to close and remove.</param>
        /// <returns>True if successfully closed and removed; otherwise false.</returns>
        public bool CloseDocument(CKDocument doc)
        {
            if (doc == null || !_documents.Contains(doc))
                return false;

            try
            {
                doc.Dispose();
                _documents.Remove(doc);
                Log.Information("Document closed and removed from tracking: {Path}", doc.FullPath);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to close document: {Message}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Unregisters a document that has been disposed or closed externally.
        /// </summary>
        /// <param name="doc">The document to unregister.</param>
        /// <returns>True if the document was removed; otherwise false.</returns>
        public bool UntrackDocument(CKDocument doc)
        {
            return _documents.Remove(doc);
        }

        /// <summary>
        /// Closes and disposes of this application and all associated documents.
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        private bool disposedValue;

        /// <summary>
        /// Internal dispose logic.
        /// </summary>
        /// <param name="disposing">True if called from Dispose(), false if from finalizer.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Log.Information("Disposing CKApplication and closing documents.");

                    foreach (var doc in _documents.ToArray())
                    {
                        try
                        {
                            if (doc.IsOrphan)
                            {
                                Log.Warning("Skipping orphaned document: {Path}", doc.FullPath);
                            }
                            else
                            {
                                doc.Dispose();
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Error disposing document: {Message}", ex.Message);
                        }
                    }

                    _documents.Clear();

                    try
                    {
                        _wordApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("Word application failed to quit cleanly: {Message}", ex.Message);
                    }
                }

                disposedValue = true;
            }
        }
    }
}
