using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Diagnostics;
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

        [Obsolete]
        public CKRange(IDOMObject parent)
        {

            _COMRange = null;
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            _isDirty = true;
        }

        #endregion

        #region Public Properties


        private CKCells _cells_1 = default;

        /// <summary>
        /// 1 based list.
        /// </summary>
        public CKCells Cells
        {
            get
            {
                if (IsDirty || _cells_1 is null)
                {
                    var COMCells_1 = COMRange.Cells;
                    var ckCells_1 = new Base1List<CKCell>();
                    foreach (Word.Cell cell in COMCells_1)
                    {
                        var cellRef = new CKCellRef(cell.RowIndex,
                            cell.ColumnIndex,
                            new RangeSnapshot(cell.Range),
                            Tables.ItemOf(cell),
                            this);
                        //each cell calls to it's own table since table is arbitrary.
                        ckCells_1.Add(new CKCell(cellRef));
                    }

                    _cells_1 = new CKCells(ckCells_1, this);
                }
                return _cells_1;
            }
        }

        /// <summary>
        /// Attempts to find the next occurrence of the specified text using Word's Find functionality.
        /// </summary>
        /// <param name="text">The text to search for.</param>
        /// <param name="matchWildcards">Whether to enable wildcard matching.</param>
        /// <param name="matchCase">Whether to perform a case-sensitive search.</param>
        /// <param name="matchWholeWord">Whether to match the whole word only.</param>
        /// <returns>
        /// A new <see cref="CKRange"/> containing the match if found; otherwise, <c>null</c>.
        /// </returns>
        /// <remarks>
        /// Version: CK2.00.01.0034
        /// </remarks>
        public CKRange TryFindNext(string text, bool matchWildcards = false, bool matchCase = false, bool matchWholeWord = false)
        {
            if (string.IsNullOrEmpty(text)) return null;

            Word.Find finder = COMRange.Find;
            finder.ClearFormatting();
            finder.Text = text;
            finder.MatchWildcards = matchWildcards;
            finder.MatchCase = matchCase;
            finder.MatchWholeWord = matchWholeWord;

            bool found = finder.Execute();

            if (found)
            {
                return new CKRange(COMRange.Duplicate, Parent);
            }

            if (Debugger.IsAttached) Debugger.Break();
            return null;
        }


        /// <summary>
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        [Obsolete]

        public Word.Range COMRange
        {
            get
            {
                if (_COMRange == null)//isDirty?
                {

                    Refresh();
                    throw new CKDebugException("Range is null");
                }
                return _COMRange;
            }
            protected set
            {
                if (_COMRange != null) throw new CKDebugException("Attempted to assign a populated Range.");
                _COMRange = value;
            }
        }

        private Word.Range _COMRange;
        /// <summary>
        /// Gets the raw text of the range as returned by Word without caching.
        /// </summary>
        public string RawText => _COMRange.Text;

        /// <summary>
        /// Gets the text of the range with caching.
        /// </summary>
        public string Text
        {

            get
            {
                this.Ping();
                if (IsDirty || _cachedText == null)
                    Refresh();

                this.Pong();
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

        private bool _isCheckingDirty = false;
        private static long _isDirtyCount = 0;
        /// <inheritdoc/>
        public virtual bool IsDirty
        {
            get
            {
                //LH.Ping($"Parent: {Parent.GetType()}", GetType());
                if (_isDirtyCount++ % 20 == 0) LH.Checkpoint($"CKRange.IsDirty count: {_isDirtyCount}");

                if (_isDirty || _isCheckingDirty)
                {
                    //this.Pong();
                    return _isDirty;
                }

                _isCheckingDirty = true;
                try
                {
                    _isDirty = _isDirty
                    || CheckDirtyFor();
                    //|| COMRange.Characters.Count != _cachedCharCount
                    //|| COMRange.Text != _cachedText;

                }
                catch (Exception ex)
                {
                    if (Debugger.IsAttached)
                    {
                        Debugger.Break();
                        throw ex;
                    }
                }
                finally
                {

                    _isCheckingDirty = false;
                }
                //this.Pong();
                return _isDirty;
            }
            protected set => _isDirty = value;
        }

        protected virtual bool CheckDirtyFor()
        {
            this.PingPong();
            return false;
        }

        private bool _checkingDirty = false;
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
        /// <remarks>overrides should always call base to ensure range cache is refreshed.</remarks>
        public void Refresh()
        {
            if (_isRefreshing) throw new CKDebugException("You are looping on Refresh. Don't do that.");
            _isRefreshing = true;
            this.Ping();
            DoRefreshThings();
            _cachedText = _COMRange.Text;
            _cachedCharCount = _COMRange.Characters.Count;
            _cachedStart = _COMRange.Start;
            _cachedEnd = _COMRange.End;
            _cachedPrettyText = CKTextHelper.Pretty(_cachedText);
            _cachedScrunchedText = CKTextHelper.Scrunch(_cachedText);
            IsDirty = false;
            _isRefreshing = false;
            this.Pong();

        }
        public bool _isRefreshing = false;

        protected virtual void DoRefreshThings()
        {
            this.PingPong();
        }

        /// <summary>
        /// Compares the text of this CKRange with the given string after scrunching (removing all whitespace).
        /// </summary>
        /// <param name="other">The string to compare.</param>
        /// <returns>True if the scrunched texts are equal; otherwise, false.</returns>
        public bool TextEquals(string other)
        {
            string myScrunched = CKTextHelper.Scrunch(_COMRange.Text);
            string otherScrunched = CKTextHelper.Scrunch(other);
            return string.Equals(myScrunched, otherScrunched, StringComparison.Ordinal);
        }

        #endregion

        public RangeSnapshot GetSnapshot()
        {
            return new RangeSnapshot(_COMRange);
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

            return RangeSnapshot.FastMatch(_COMRange, other._COMRange);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int start = IsOrphan ? _cachedStart : _COMRange.Start;
                int end = IsOrphan ? _cachedEnd : _COMRange.End;
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
            _COMRange.Delete();
        }
    }

}
