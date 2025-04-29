using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using K4os.Hash.xxHash;
using Serilog;
using System;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a wrapper around a Word document, providing DOM access to tables, sections,
    /// headers, footers, and other editable regions of the document.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0001
    /// </remarks>
    public class CKDocument : IDOMObject, IDisposable
    {
        protected readonly string _fullPath;
        protected Word.Document _comDocument;
        private CKTables _tables;
        private CKSections _sections;
        private CKParagraphs _paragraphs;



        /// <summary>
        /// Returns a duplicated COM reference to the underlying Word.Document.
        /// Caller is responsible for releasing it via <see cref="System.Runtime.InteropServices.Marshal.ReleaseComObject"/>.
        /// </summary>
        /// <returns>A new RCW to the internal Word.Document.</returns>
        /// <remarks>Version: CK2.00.01.0020</remarks>
        public Word.Document GiveMeCOMDocumentIWillOwnItAndPromiseToCleanUpAfterMyself()
        {
            if (_comDocument == null)
                throw new InvalidOperationException("Underlying COM document is not initialized.");

            return _comDocument.Application.Documents.Open(FullPath);
        }

        /// <summary>
        /// The CKApplication instance that owns and opened this document.
        /// </summary>
        public CKApplication Application { get; private set; }

        /// <summary>
        /// The full file path of the underlying document.
        /// </summary>
        public string FullPath => _fullPath;

        public string FileName => Path.GetFileName(_fullPath);

        /// <summary>
        /// Provides access to the document's tables as a CKTables collection.
        /// </summary>

        public CKTables Tables => this.PingPong(() => Content.Tables, msg: "$$$");


        /// <summary>
        /// Provides access to the document's sections as a <see cref="CKSections"/> collection.
        /// </summary>
        public CKSections Sections => Content.Sections;

        /// <inheritdoc/>
        public CKDocument Document => this;

        /// <inheritdoc/>
        public IDOMObject Parent => throw new NotSupportedException("Call Application on a CKDocument object.");

        private bool _isDirty = false;
        private bool _isCheckingDirty = false;

        /// <inheritdoc/>
        public bool IsDirty
        {
            get
            {
                if (_isDirty || _isCheckingDirty)
                    return _isDirty;

                _isCheckingDirty = true;
                try
                {
                    _isDirty =
                        _tables?.IsDirty == true ||
                        _sections?.IsDirty == true ||
                        _paragraphs?.IsDirty == true;
                }

                finally
                {
                    _isCheckingDirty = false;
                }

                return _isDirty;
            }
            protected set => _isDirty = value;
        }


        /// <inheritdoc/>
        public bool IsOrphan
        {
            get
            {
                try { _ = _comDocument.FullName; return false; }
                catch (COMException) { return true; }
                catch (Exception) { return true; }
            }
        }
        private string _logId;

        /// <summary>
        /// Gets a short unique identifier for this document instance, suitable for log tracing.
        /// </summary>
        public string LogId => _logId is null ? GenerateLogId() : _logId;

        public bool Visible
        {
            get => _comDocument.Windows[1].Visible;
            set => _comDocument.Windows[1].Visible = value;
        }
        public CKRange Content => new CKRange(_comDocument.Content, this);


        /// <summary>
        /// Provides access to the document's paragraphs as a <see cref="CKParagraphs"/> collection.
        /// </summary>
        public CKParagraphs Paragraphs
        {
            get
            {
                if (_paragraphs == null || _paragraphs.IsDirty)
                {
                    _paragraphs = new CKParagraphs(_comDocument.Paragraphs, this);
                }
                return _paragraphs;
            }
        }

        public bool KeepAlive { get; internal set; }
        [Obsolete]
        public Word.Window ActiveWindow => _comDocument.ActiveWindow;

        public bool Saved
        {
            get => _comDocument.Saved;
            set => _comDocument.Saved = value;
        }
        public bool ReadOnlyRecommended
        {
            get => _comDocument.ReadOnlyRecommended;
            set => _comDocument.ReadOnlyRecommended = value;
        }
        /// <summary>
        /// Gets or sets whether this document is marked as Final (read-only UI state).
        /// Only valid when this document is the active document and Word has at least one document open.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.01.0003
        /// </remarks>
        public bool Final
        {
            get
            {
                try
                {
                    if (_comDocument.Application.Documents.Count == 0)
                    {
                        Log.Debug("No documents open. Cannot get Final property.");
                        return false;
                    }

                    if (_comDocument.Application.ActiveDocument == _comDocument)
                    {
                        return _comDocument.Final;
                    }

                    Log.Debug("This document is not the active document. Cannot get Final property.");
                    return false;
                }
                catch (COMException ex)
                {
                    Log.Warning("Failed to get Final property: {Message}", ex.Message);
                    return false;
                }
            }
            set
            {
                try
                {
                    if (_comDocument.Application.Documents.Count == 0)
                    {
                        Log.Debug("No documents open. Cannot set Final property.");
                        return;
                    }
#if DEBUG
                    if (_comDocument.Application.ActiveDocument == _comDocument)
                    {
                        _comDocument.Final = value;
                    }
#endif
                    else
                    {
                        Log.Debug("This document is not the active document. Cannot set Final property.");
                    }
                }
                catch (COMException ex)
                {
                    Log.Warning("Failed to set Final property: {Message}", ex.Message);
                }
            }
        }



        public void Activate()
        {

            _comDocument.Activate();

            // Force layout pass
            _comDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            _comDocument.ActiveWindow.View.Zoom.Percentage = 100;

            // Sleep briefly to let Word render
            System.Threading.Thread.Sleep(250);
        }
        private string GenerateLogId()
        {
            try
            {
                var seed = $"{DateTime.UtcNow.Ticks}|{Guid.NewGuid()}|DOC|{_fullPath}";
                var bytes = System.Text.Encoding.UTF8.GetBytes(seed);
                ulong hash = XXH64.DigestOf(bytes);
                return hash.ToString("X8");
            }
            catch
            {
                return "UNKNOWN";
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKDocument"/> class from an existing Word.Document.
        /// </summary>
        /// <param name="wordDoc">The COM document to wrap.</param>
        /// <param name="app">The owning CKApplication (optional).</param>
        public CKDocument(Word.Document wordDoc, CKApplication app = null)
        {
            _comDocument = wordDoc ?? throw new ArgumentNullException(nameof(wordDoc));
            _fullPath = _comDocument.FullName;
            Application = app;
        }

        public Tracer Tracer = new Tracer(typeof(CKDocument));
        private bool disposedValue;

        /// <summary>
        /// Deletes the section at the specified 1-based index, including its section break.
        /// </summary>
        /// <param name="sectionIndex">The 1-based index of the section to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if index is outside valid range.</exception>
        /// <summary>
        /// Deletes the specified section in a Word document, including all its content
        /// and the section break that follows it.
        /// </summary>
        /// <param name="doc">The Word document containing the section.</param>
        /// <param name="sectionIndex">The 1-based index of the section to delete.</param>
        public void DeleteSection(int sectionIndex)
        {
            var sections = _comDocument.Sections;
            // Word uses 1-based indexing for sections.
            if (sectionIndex < 1 || sectionIndex > sections.Count)
                throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Invalid section index.");

            // Get the section to delete.
            Word.Section section = sections[sectionIndex];

            // Get the full range of the section. This includes the section break that follows it.
            Word.Range sectionRange = section.Range;

            // If this is the *last* section, the range includes the final paragraph mark of the document.
            // Trying to delete that will make Word sad (and by sad, I mean unstable or crashy).
            if (sectionIndex == sections.Count)
            {
                // So we shrink the range by 1 character to preserve the final paragraph mark.
                sectionRange.End -= 1;
            }

            // Delete the section and everything in it, including the section break *after* it.
            // The section break *before* it (if any) is NOT deleted, so the document remains well-formed.
            sectionRange.Delete();

            // Done. Word will automatically renumber the remaining sections.
        }

        /// <summary>
        /// Gets the primary footer range of the first section.
        /// </summary>
        public Word.Range GetFooterRange() => _comDocument.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

        /// <summary>
        /// Gets the primary header range of the first section.
        /// </summary>
        public Word.Range GetHeaderRange() => _comDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

        /// <summary>
        /// Copies the footer from this document into the target document.
        /// </summary>
        /// <param name="targetDocument">The document to receive the footer content.</param>
        public void CopyFooterTo(CKDocument targetDocument)
        {
            if (targetDocument == null)
                throw new ArgumentNullException(nameof(targetDocument));
            targetDocument.GetFooterRange().FormattedText = GetFooterRange().FormattedText;
        }

        /// <summary>
        /// Copies the header from this document into the target document.
        /// </summary>
        /// <param name="targetDocument">The document to receive the header content.</param>
        public void CopyHeaderTo(CKDocument targetDocument)
        {
            if (targetDocument == null)
                throw new ArgumentNullException(nameof(targetDocument));
            targetDocument.GetHeaderRange().FormattedText = GetHeaderRange().FormattedText;
        }

        /// <summary>
        /// Copies both header and footer from this document to the target document.
        /// </summary>
        /// <param name="targetDocument">The document to receive the content.</param>
        public void CopyHeaderAndFooterTo(CKDocument targetDocument)
        {
            CopyHeaderTo(targetDocument);
            CopyFooterTo(targetDocument);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Application?.UntrackDocument(this); //WHY application null bypass?
                }
                disposedValue = true;
            }
        }

        /// <summary>
        /// Gets a wrapper around the entire document range.
        /// </summary>
        public CKRange Range() => Range(_comDocument.Range());

        /// <summary>
        /// Gets a wrapper around a specific sub-range of the document.
        /// </summary>
        /// <param name="start">Start position (inclusive).</param>
        /// <param name="end">End position (inclusive).</param>
        public CKRange Range(int start, int end) => Range(_comDocument.Range(start, end));

        /// <summary>
        /// Wraps an existing Word.Range as a CKRange.
        /// </summary>
        /// <param name="range">The Word.Range to wrap.</param>
        public CKRange Range(Word.Range range) => new CKRange(range, this);

        public void Dispose()
        {
            if (!KeepAlive)
            {
                Dispose(disposing: true);
                GC.SuppressFinalize(this);
            }
        }

        internal void RemoveDocumentInformation(Word.WdRemoveDocInfoType wdRDIDocumentProperties)
        {
            _comDocument.RemoveDocumentInformation(wdRDIDocumentProperties);
        }

        /// <summary>
        /// Check if object has a refernce to the same Word.Document.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Matches(object other)
        {
            LH.Debug("Tracker[!sd] CKDocument::");

            if (_comDocument == null || other == null) return false;

            if (ReferenceEquals(this, other)) return true;
            if (Matches(other as Word.Document)) return true;
            if (other is Word.Range wdRange && Matches(wdRange.Document)) return true;
            if (other is IDOMObject iDom && Equals(iDom.Document)) return true;

            //TODO other matches?
            return false;
        }
        public bool Matches(Word.Document other)
        {
            LH.Debug("Tracker[!sd]");
            if (ReferenceEquals(_comDocument, other)) return true;

            try
            {
                var thisRange = _comDocument?.Content;
                var otherRange = other?.Content;

                if (thisRange == null || otherRange == null) return false;

                return RangeSnapshot.SlowMatch(thisRange, otherRange);
            }
            catch
            {
                return false;
            }
        }
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (!(obj is CKDocument other)) return false;

            try
            {
                var thisRange = _comDocument?.Content;
                var otherRange = other._comDocument?.Content;

                if (thisRange == null || otherRange == null) return false;

                return RangeSnapshot.SlowMatch(thisRange, otherRange);
            }
            catch
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            try
            {
                var range = _comDocument?.Content;
                if (range == null) return 0;

                var hash = 23 + _comDocument.GetHashCode();
                return hash;
            }
            catch
            {
                return 0;
            }
        }

        public bool EnsureLayoutReady() => EnsureLayoutReady(_comDocument);

        /// <summary>
        /// Forces Word to complete document layout and rendering to ensure safe access to Ranges, Tables, and other elements.
        /// </summary>
        /// <remarks>Version: CK2.00.01.0021</remarks>
        public static bool EnsureLayoutReady(Word.Document comDocument)
        {
            try
            {
                if (comDocument == null || comDocument.Application == null)
                {
                    Log.Warning("EnsureLayoutReady called, but document or application is null.");
                    return false;
                }

                var window = comDocument.ActiveWindow;
                if (window != null)
                {
                    window.View.Type = Word.WdViewType.wdPrintView;
                    window.View.Zoom.Percentage = 100;
                }

                System.Threading.Thread.Sleep(250);
                comDocument.Repaginate();

                Log.Debug("EnsureLayoutReady completed successfully for document: {FileName}",
                    Path.GetFileName(comDocument.FullName));
                _EnsureLayoutReady_depth = 0; // ✅ reset depth on success
                return true;
            }
            catch (COMException ex)
            {
                Log.Error("EnsureLayoutReady failed: {Message}", ex.Message);
            }
            catch (Exception ex)
            {
                Log.Error("Unexpected error during EnsureLayoutReady: {Message}", ex.Message);
            }

            if (_EnsureLayoutReady_depth++ < 10)
            {
                Log.Debug("...Waiting for document layout. Retry #{_EnsureLayoutReady_depth}...");
                return EnsureLayoutReady(comDocument);
            }
            _EnsureLayoutReady_depth = 0; // ✅ reset depth on fail if caller wants to try again
            return false;
        }
        //address if this is too restrictive, such as too many documents ensuring at once.
        private static int _EnsureLayoutReady_depth;

    }
}
