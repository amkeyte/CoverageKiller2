using CoverageKiller2.DOM.Tables;
using Serilog;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a simple wrapper for the Word.Range object.
    /// Provides caching of text and boundary values for robust equality and hash code calculations,
    /// even if the underlying COM object becomes orphaned.
    /// </summary>
    public class CKRange : IDOMObject, IDisposable
    {
        CKApplication IDOMObject.Application => Parent.Application;

        #region Fields & Caching

        private string _cachedText;
        private string _cachedPrettyText;
        private string _cachedScrunchedText;
        private int _cachedCharCount;
        private int _cachedStart;
        private int _cachedEnd;
        private bool _isDirty = false;
        private bool _isOrphan = false;
        private bool disposedValue;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <param name="parent">Optional parent DOM object; if not provided, will be looked up via CKDocuments.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> parameter is null.</exception>
        public CKRange(Word.Range range, IDOMObject parent)
        {

            COMRange = range ?? throw new ArgumentNullException(nameof(range));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            _cachedCharCount = COMRange.Characters.Count;
            _cachedText = COMRange.Text;
            // Initialize cached boundary values.
            _cachedStart = COMRange.Start;
            _cachedEnd = COMRange.End;
        }

        #endregion

        #region Public Properties


        private CKCells _cells = default;
        public CKCells Cells
        {
            get
            {
                if (IsDirty || _cells is null)
                {
                    var COMCells = COMRange.Cells;
                    var cellsRefList = new List<CKCellRef>();
                    foreach (Word.Cell cell in COMCells)
                    {
                        var cellRef = new CKCellRef(cell.RowIndex,
                            cell.ColumnIndex,
                            new RangeSnapshot(cell.Range),
                            this);
                        cellsRefList.Add(cellRef);
                    }
                    var cellsRef = new CellsRef(cellsRefList, this);
                    _cells = new CKCells(COMCells, cellsRef);
                }
                return _cells;
            }
        }

        /// <summary>
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        public Word.Range COMRange { get; private set; }

        /// <summary>
        /// Gets the raw text of the range as returned by Word.
        /// </summary>
        public string RawText => COMRange.Text;

        /// <summary>
        /// Gets the text of the range (raw text).
        /// </summary>
        public string Text
        {
            get
            {
                if (IsDirty || _cachedText == null)
                    Refresh();
                return _cachedText;
            }
            set
            {
                IsDirty = true;
                COMRange.Text = value;
            }
        }



        /// <summary>
        /// Gets a "pretty" version of the range's text.
        /// This version replaces cell markers with tabs, preserves Windows-style newlines,
        /// and removes extraneous control characters.
        /// </summary>
        public string PrettyText
        {
            get
            {
                if (IsDirty || _cachedPrettyText == null)
                    Refresh();
                return _cachedPrettyText;
            }
        }

        /// <summary>
        /// Gets the scrunched version of the range's text, i.e. all whitespace removed,
        /// for reliable comparisons.
        /// </summary>
        public string ScrunchedText
        {
            get
            {
                if (IsDirty || _cachedScrunchedText == null)
                    Refresh();
                return _cachedScrunchedText;
            }
        }

        /// <summary>
        /// Gets the starting position of the range.
        /// </summary>
        public int Start => COMRange.Start;

        /// <summary>
        /// Gets the ending position of the range.
        /// </summary>
        public int End => COMRange.End;

        /// <summary>
        /// Gets a value indicating whether this CKRange is dirty.
        /// It becomes dirty when the underlying COMRange has changed.
        /// </summary>
        public virtual bool IsDirty
        {
            get
            {
                _isDirty = _isDirty
                    || COMRange.Characters.Count != _cachedCharCount
                    || COMRange.Text != _cachedText;
                return _isDirty;
            }
            set => _isDirty = value;
        }

        /// <summary>
        /// Gets a value indicating whether this CKRange is orphaned,
        /// i.e. its underlying COMRange is no longer valid.
        /// </summary>
        public bool IsOrphan
        {
            get
            {
                if (_isOrphan)
                    return true;
                try
                {
                    // Access a property to check if the COM object is valid.
                    _ = COMRange.Text;
                    _isOrphan = false;
                }
                catch (COMException ex)
                {
                    Log.Error("Attempt to access orphan range failed.", ex);
                    _isOrphan = true;
                }
                catch (Exception)
                {
                    _isOrphan = true;
                }
                return _isOrphan;
            }
        }

        /// <summary>
        /// Gets the parent DOM object.
        /// </summary>
        public IDOMObject Parent { get; protected set; }

        /// <summary>
        /// Gets the document associated with this CKRange.
        /// </summary>
        public CKDocument Document => Parent.Document;

        /// <summary>
        /// Gets the Word application managing the document.
        /// </summary>
        public CKApplication Application => Document.Application;

        /// <summary>
        /// Gets the sections contained in the range.
        /// </summary>
        public CKSections Sections => new CKSections(COMRange.Sections, this);

        /// <summary>
        /// Gets the paragraphs contained in the range.
        /// </summary>
        public CKParagraphs Paragraphs => new CKParagraphs(COMRange.Paragraphs, this);

        /// <summary>
        /// Gets the tables contained in the range.
        /// </summary>
        public CKTables Tables => new CKTables(COMRange.Tables, this);
        /// <summary>
        /// Gets or sets the formatted text for this range. Setting this value replaces the contents and formatting of the range.
        /// </summary>
        public CKRange FormattedText
        {
            get
            {
                if (COMRange == null) throw new InvalidOperationException("COMRange is null.");
                var formatted = COMRange.FormattedText;
                return new CKRange(formatted, Parent);
            }
            set
            {
                if (COMRange == null) throw new InvalidOperationException("COMRange is null.");
                if (value?.COMRange == null) throw new ArgumentNullException(nameof(value));
                COMRange.FormattedText = value.COMRange;
            }
        }
        /// <summary>
        /// Returns a new CKRange collapsed to the end of this range.
        /// </summary>
        /// <returns>A new CKRange positioned at the end of this range.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0015
        /// </remarks>
        public CKRange CollapseToEnd()
        {
            int end = COMRange.End;
            var docRange = Document.Range();
            int max = Math.Max(0, docRange.End - 1);

            // Clamp end to valid range
            if (end > max) end = max;
            if (end < 0) end = 0;

            Word.Range collapsed;
            try
            {
                collapsed = Document.Range(end, end).COMRange;
            }
            catch
            {
                collapsed = docRange.COMRange; // fallback: entire document
            }

            return new CKRange(collapsed, Document);
        }

        /// <summary>
        /// Returns a new CKRange collapsed to the start of this range.
        /// </summary>
        /// <returns>A new CKRange positioned at the start of this range.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0016
        /// </remarks>
        public CKRange CollapseToStart()
        {
            int start = COMRange.Start;
            var docRange = Document.Range();
            int max = Math.Max(0, docRange.End - 1);

            if (start > max) start = max;
            if (start < 0) start = 0;

            Word.Range collapsed;
            try
            {
                collapsed = Document.Range(start, start).COMRange;
            }
            catch
            {
                collapsed = docRange.COMRange; // fallback: entire document
            }

            return new CKRange(collapsed, Document);
        }

        /// <summary>
        /// Gets the cells contained in the range.
        /// </summary>
        //public CKCells Cells => new CKCellsLinear(this);

        #endregion

        #region Public Methods

        /// <summary>
        /// Updates cached text values from the underlying COMRange and resets the dirty flag.
        /// </summary>
        public void Refresh()
        {
            _cachedText = COMRange.Text;
            _cachedCharCount = COMRange.Characters.Count;
            _cachedStart = COMRange.Start;
            _cachedEnd = COMRange.End;
            _cachedPrettyText = CKTextHelper.Pretty(_cachedText);
            _cachedScrunchedText = CKTextHelper.Scrunch(_cachedText);
            IsDirty = false;
        }

        /// <summary>
        /// Compares the text of this CKRange with the given string after scrunching (removing all whitespace).
        /// </summary>
        /// <param name="other">The string to compare.</param>
        /// <returns>True if the scrunched texts are equal; otherwise, false.</returns>
        public bool TextEquals(string other)
        {
            string myScrunched = CKTextHelper.Scrunch(COMRange.Text);
            string otherScrunched = CKTextHelper.Scrunch(other);
            return string.Equals(myScrunched, otherScrunched, StringComparison.Ordinal);
        }

        #endregion

        public RangeSnapshot GetSnapshot()
        {
            return new RangeSnapshot(COMRange);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;
            if (obj == null || !(obj is CKRange))
                return false;
            return Equals((CKRange)obj);
        }

        public bool Equals(CKRange other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other == null)
                return false;

            return RangeSnapshot.FastMatch(COMRange, other.COMRange);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int start = IsOrphan ? _cachedStart : COMRange.Start;
                int end = IsOrphan ? _cachedEnd : COMRange.End;
                int hash = 17;
                hash = hash * 23 + start.GetHashCode();
                hash = hash * 23 + end.GetHashCode();
                return hash;
            }
        }

        public static bool operator ==(CKRange left, CKRange right)
        {
            if (ReferenceEquals(left, right))
                return true;
            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
                return false;
            return left.Equals(right);
        }

        public static bool operator !=(CKRange left, CKRange right)
        {
            return !(left == right);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~CKRange()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        internal void Delete()
        {
            COMRange.Delete();
        }
    }

}
