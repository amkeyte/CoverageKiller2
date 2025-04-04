using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using System;
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
        private readonly string _fullPath;
        private Word.Document _comDocument;

        /// <summary>
        /// The CKApplication instance that owns and opened this document.
        /// </summary>
        public CKApplication Application { get; private set; }

        /// <summary>
        /// The full file path of the underlying document.
        /// </summary>
        public string FullPath => _fullPath;

        /// <summary>
        /// Provides access to the document's tables as a CKTables collection.
        /// </summary>
        public CKTables Tables => new CKTables(Range());

        /// <summary>
        /// Provides access to the document's sections.
        /// </summary>
        public CKSections Sections => new CKSections(Range());

        /// <inheritdoc/>
        public CKDocument Document => this;

        /// <inheritdoc/>
        public IDOMObject Parent => throw new NotSupportedException("Call Application on a CKDocument object.");

        /// <inheritdoc/>
        public bool IsDirty => throw new NotImplementedException();

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
        public void DeleteSection(int sectionIndex)
        {
            Tracer.Log("Deleting Section", new DataPoints().Add(nameof(sectionIndex), sectionIndex));
            if (sectionIndex < 1 || sectionIndex > _comDocument.Sections.Count)
                throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Section index is out of range.");
            Word.Section sectionToDelete = _comDocument.Sections[sectionIndex];
            Word.Range extendedRange = _comDocument.Range(sectionToDelete.Range.Start - 1, sectionToDelete.Range.End);
            extendedRange.Delete();
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
                    Application?.UntrackDocument(this);
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
        public CKRange Range(Word.Range range) => new CKRange(range);

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
