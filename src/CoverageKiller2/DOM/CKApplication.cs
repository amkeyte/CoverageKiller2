using CoverageKiller2.Logging;
using Microsoft.Office.Interop.Word;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a Word application instance responsible for creating, managing,
    /// and closing documents in a controlled environment.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0009
    /// </remarks>
    public class CKApplication : IDisposable
    {
        private readonly Application _wordApp;
        private readonly List<CKDocument> _documents = new List<CKDocument>();
        private readonly Dictionary<CKDocument, Document> _comDocs = new Dictionary<CKDocument, Document>();
        private readonly string _PID;
        private bool disposedValue;

        /// <summary>
        /// Whether this CKApplication instance is responsible for disposing the Word instance.
        /// </summary>
        public bool IsOwned { get; private set; }

        /// <summary>
        /// Gets the raw Word application instance.
        /// </summary>
        public Application WordApp => _wordApp;

        /// <summary>
        /// Indicates whether this application is the VSTO ThisAddIn instance.
        /// </summary>
        public bool IsAddIn => ReferenceEquals(_wordApp, Globals.ThisAddIn?.Application);

        /// <summary>
        /// Gets the process ID of the associated Word instance.
        /// </summary>
        public string PID => _PID;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKApplication"/> class.
        /// </summary>
        /// <param name="wordApp">The Word application to wrap.</param>
        /// <param name="pid">The process ID of the Word instance.</param>
        /// <param name="isOwned">Whether CKOffice is responsible for cleanup.</param>
        public CKApplication(Application wordApp, int pid, bool isOwned = true)
        {
            LH.Ping(GetType());
            _wordApp = wordApp ?? throw new ArgumentNullException(nameof(wordApp));
            IsOwned = isOwned;
            _PID = pid.ToString();

            Log.Verbose("CKApplication ctor success [{PID}] (Owned={IsOwned})", _PID, IsOwned);
            LH.Pong(GetType());
        }

        /// <summary>
        /// Gets all CKDocuments currently tracked by this application.
        /// </summary>
        public IReadOnlyList<CKDocument> Documents => _documents.AsReadOnly();

        /// <summary>
        /// Opens a document from disk and wraps it in a CKDocument.
        /// </summary>
        public CKDocument GetDocument(string fullPath, bool visible = false)
        {
            LH.Ping(GetType());
            if (string.IsNullOrWhiteSpace(fullPath))
                throw new ArgumentException("Invalid file path.", nameof(fullPath));

            try
            {
                Log.Information("Opening document from path: {Path}", fullPath);
                var comDoc = _wordApp.Documents.Open(
                    FileName: fullPath,
                    ReadOnly: true,
                    Visible: visible
                );

                var ckDoc = new CKDocument(comDoc, this);
                _documents.Add(ckDoc);
                _comDocs[ckDoc] = comDoc;

                Log.Information("Document opened and tracked: {fileName}", ckDoc.FileName);
                LH.Pong(GetType());
                return ckDoc;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to open document {Path}: {Message}", fullPath, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Closes and removes the specified document.
        /// </summary>
        public bool CloseDocument(CKDocument doc, bool force = false)
        {
            if (doc == null || !_documents.Contains(doc))
                return false;

            try
            {
                if (force && _comDocs.TryGetValue(doc, out var comDoc))
                {
                    try
                    {
                        comDoc.Close(SaveChanges: false);
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("Force-close failed on COM document: {Message}", ex.Message);
                    }
                }
                else
                {
                    if (Debugger.IsAttached) Debugger.Break();
                }

                doc.Dispose(); // Calls UntrackDocument internally
                _documents.Remove(doc);
                _comDocs.Remove(doc);

                Log.Information("Document closed and removed from tracking: {FileName}", doc.FileName);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to close document: {Message}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Stops tracking a disposed document.
        /// </summary>
        public bool UntrackDocument(CKDocument doc)
        {
            _comDocs.Remove(doc);
            return _documents.Remove(doc);
        }

        /// <summary>
        /// Disposes the application and optionally quits Word.
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Internal dispose logic. Quits Word only if this is an owned instance.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Log.Information($"Disposing CKApplication({PID}) and closing documents.");

                    foreach (var doc in _documents.ToArray())
                    {
                        try { CloseDocument(doc, force: true); }
                        catch (Exception ex)
                        {
                            Log.Error("Error disposing document: {Message}", ex.Message);
                        }
                    }

                    _documents.Clear();
                    _comDocs.Clear();

                    if (!IsOwned)
                    {
                        Log.Information("CKApplication({PID}) is not owned; skipping WordApp.Quit().", PID);
                    }
                    else
                    {
                        try
                        {
                            _wordApp.Quit();
                        }
                        catch (Exception ex)
                        {
                            Log.Warning("Word application failed to quit cleanly: {Message}", ex.Message);
                        }
                    }
                }

                disposedValue = true;
            }
        }
    }
}
