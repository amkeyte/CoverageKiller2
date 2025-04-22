using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a Word application instance responsible for creating, managing,
    /// and closing documents in a controlled environment.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0009
    /// </remarks>
    public partial class CKApplication : IDisposable
    {
        private readonly Word.Application _wordApp;
        private List<CKDocument> _documents = new List<CKDocument>();
        private readonly Dictionary<CKDocument, Word.Document> _comDocs = new Dictionary<CKDocument, Word.Document>();
        private readonly string _PID;
        private bool disposedValue;

        /// <summary>
        /// Whether this CKApplication instance is responsible for disposing the Word instance.
        /// </summary>
        public bool IsOwned { get; private set; }

        /// <summary>
        /// Gets the raw Word application instance.
        /// </summary>
        public Word.Application WordApp => _wordApp;

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
        public CKApplication(Word.Application wordApp, int pid, bool isOwned = true)
        {
            this.Ping(msg: "$$$");
            _wordApp = wordApp ?? throw new ArgumentNullException(nameof(wordApp));
            IsOwned = isOwned;
            _PID = pid.ToString();

            Log.Verbose("CKApplication ctor success [{PID}] (Owned={IsOwned})", _PID, IsOwned);
            this.Pong();
        }

        /// <summary>
        /// Gets all CKDocuments currently tracked by this application.
        /// </summary>
        public IReadOnlyList<CKDocument> Documents => _documents.AsReadOnly();



        /// <summary>
        /// Opens a document from disk and wraps it in a CKDocument.
        /// </summary>
        public CKDocument GetDocument(string fullPath, bool visible = false, bool createIfNotFound = false)
        {
            this.Ping(msg: "$$$");
            if (string.IsNullOrWhiteSpace(fullPath))
                throw new ArgumentException("Invalid file path.", nameof(fullPath));

            try
            {
                //Word.Document comDoc = default;
                if (createIfNotFound && !File.Exists(fullPath))
                    _wordApp.Documents.Add(Visible: visible).SaveAs2(FileName: fullPath);

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
                this.Pong();
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
            if (doc.KeepAlive)
            {
                Log.Warning($"Document {doc.FileName} was requested to close, but KeepAlive is true.");
                return false;
            }
            try
            {
                if (force && _comDocs.TryGetValue(doc, out var comDoc))
                {
                    try
                    {
                        if (doc.KeepAlive)
                        {
                            doc.KeepAlive = false;
                            Log.Warning($"Document {doc.FileName} was requested to close." +
                                $" KeepAlive is true, but was overriden by Force-close.");
                        }
                        comDoc.Saved = true;
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
                //_documents.Remove(doc);
                //_comDocs.Remove(doc);

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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_crashing)
                        Log.Warning($"Disposing CKApplication({PID}) due to crash. Forcing all documents closed.");
                    else
                        Log.Information($"Disposing CKApplication({PID}) and closing documents.");

                    foreach (var doc in _documents.ToArray())
                    {
                        if (!_crashing && doc.KeepAlive) continue;

                        try
                        {
                            CloseDocument(doc, force: true);
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Error disposing document: {Message}", ex.Message);
                        }
                    }

                    if (!_crashing && HasKeepOpenDocuments)
                    {
                        // Do not clear collections — documents are still in use
                    }
                    else
                    {
                        _documents.Clear();
                        _comDocs.Clear();
                    }

                    bool blockDispose = !_crashing && (!IsOwned || HasKeepOpenDocuments);

                    if (blockDispose)
                    {
                        if (!IsOwned)
                        {
                            Log.Information("CKApplication({PID}) is not owned; skipping WordApp.Quit().", PID);
                        }
                        else if (HasKeepOpenDocuments)
                        {
                            Log.Information("CKApplication({PID}) has KeepOpen documents; skipping WordApp.Quit().", PID);
                        }
                    }
                    else
                    {
                        Log.Information("CKApplication({PID}) proceeding to quit WordApp. Reason: {Reason}",
                            PID,
                            _crashing ? "Crash override" : "Clean shutdown with no KeepAlive documents");

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

        private bool _crashing = false;
        internal void Crash(Type callerType, string callerMember)
        {
            _crashing = true;
            IsOwned = true;

            Log.Error($"Crashing {nameof(CKApplication)}. Source: {callerType.Name}.{callerMember}");
            for (int docIndex = 0; docIndex < _documents.Count; docIndex++)
            {
                WithSuppressedAlerts(() =>
                {
                    CloseDocument(Documents[docIndex], force: true);
                });

            }
        }

        public bool Visible
        {
            get => _wordApp.Visible;
            set => _wordApp.Visible = value;
        }
        public bool HasKeepOpenDocuments => _documents.Any(doc => doc.KeepAlive);

        /// <summary>
        /// Gets the currently active CKDocument in the Word application, if any.
        /// Returns null if no document is active or the ActiveDocument call fails.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.01.0001
        /// </remarks>
        public CKDocument ActiveDocument
        {
            get
            {
                try
                {
                    var activeComDoc = _wordApp.ActiveDocument;
                    return GetDocument(activeComDoc.FullName);
                }
                catch (COMException)
                {
                    Log.Debug("ActiveDocument is unavailable: no document is currently active.");
                    return null;
                }
            }
        }
    }

    public partial class CKApplication
    {
        private const string DefaultTemplatePath = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\_shadowDocumentTemplate_.docx";

        /// <summary>
        /// Executes the given action with Word alerts and macro security suppressed.
        /// </summary>
        /// <param name="action">The action to run with suppressed alerts.</param>
        public void WithSuppressedAlerts(Action action)
        {
            WithSuppressedAlerts<object>(() => { action(); return null; });
        }
        /// <summary>
        /// Executes a function with Word alerts and macro security suppressed, returning a result.
        /// </summary>
        /// <typeparam name="T">The return type.</typeparam>
        /// <param name="func">The function to run with suppressed alerts.</param>
        /// <returns>The result of the function.</returns>
        public T WithSuppressedAlerts<T>(Func<T> func)
        {
            this.Ping();
            var originalAlerts = WordApp.DisplayAlerts;
            var originalSecurity = WordApp.AutomationSecurity;

            try
            {
                WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                WordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                return func();
            }
            finally
            {
                WordApp.DisplayAlerts = originalAlerts;
                WordApp.AutomationSecurity = originalSecurity;
            }
            this.Pong();
        }

        /// <summary>
        /// Creates a temporary CKDocument based on the given file path.
        /// </summary>
        /// <param name="fromFile">The source file to clone. If empty, uses the default template.</param>
        /// <returns>A new CKDocument instance opened in this application.</returns>
        public CKDocument GetTempDocument(string fromFile = "")
        {
            this.Ping(msg: "$$$");
            fromFile = string.IsNullOrWhiteSpace(fromFile) ? DefaultTemplatePath : fromFile;
            if (!File.Exists(fromFile)) throw new FileNotFoundException("Template file not found.", fromFile);

            var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + Path.GetExtension(fromFile));
            File.Copy(fromFile, tempPath);

            CKDocument doc = null;
            WithSuppressedAlerts(() => doc = GetDocument(tempPath, visible: false));
            doc.Saved = true;
            doc.ReadOnlyRecommended = false;
            doc.Final = false;
            doc.RemoveDocumentInformation(Word.WdRemoveDocInfoType.wdRDIDocumentProperties);
            this.Pong();
            return doc;
        }

        /// <summary>
        /// Creates a ShadowWorkspace using a hidden, disposable CKDocument.
        /// </summary>
        /// <returns>A new ShadowWorkspace instance.</returns>
        public ShadowWorkspace GetShadowWorkspace(bool keepOpen = false)
        {
            this.Ping(msg: "$$$");
            var doc = GetTempDocument();
            var workspace = new ShadowWorkspace(doc, this, keepOpen);
            this.Pong();
            return workspace;
        }
    }
}
