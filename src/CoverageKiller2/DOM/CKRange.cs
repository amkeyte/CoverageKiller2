using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.IO;
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
        private int _cachedStart;
        private int _cachedEnd;
        private bool _isDirty = false;
        private bool _isOrphan = false;
        private bool _disposedValue;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <param name="parent">Parent DOM object; if not provided, will be looked up via CKDocuments.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> parameter is null.</exception>
        public CKRange(Word.Range range, IDOMObject parent, bool deferCom = false) : this(parent, deferCom)
        {
            this.Ping(msg: "$$$");
            COMRange = range ?? throw new ArgumentNullException(nameof(range)); //document match done in here
            IsCOMDeferred = false;
            var msg = $"_COMRange:[{Path.GetFileName(_COMRange.Document.FullName)}::{new RangeSnapshot(_COMRange).FastHash}]" +
                $"CKRamge:[{Document.FileName}::{Snapshot.FastHash}";

            this.Pong(msg: msg);
        }

        public CKRange(IDOMObject parent, bool deferCom = true)
        {
            this.Ping(msg: "$$$");
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            _COMRange = null;
            IsCOMDeferred = deferCom;
            IsDirty = IsCOMDeferred;
            this.Pong(msg: $"{Document.FileName}::[COMRANGE NOT ASSIGNED]");
        }

        #endregion

        #region Public Properties


        private CKCells _cached_cells_1 = default;

        /// <summary>
        /// 1 based list.
        /// </summary>
        //public CKCells Cells3
        //{
        //    get
        //    {
        //        if (IsDirty || _cached_cells_1 is null)
        //        {
        //            var COMCells_1 = COMRange.Cells;
        //            var ckCells_1 = new Base1List<CKCell>();
        //            foreach (Word.Cell cell in COMCells_1)
        //            {
        //                var cellRef = new CKCellRef(cell.RowIndex,
        //                    cell.ColumnIndex,
        //                    new RangeSnapshot(cell.Range),
        //                    Tables.ItemOf(cell),
        //                    this);

        //                ckCells_1.Add(new CKCell(cellRef));
        //            }

        //            _cached_cells_1 = new CKCells(ckCells_1, this);
        //        }
        //        return _cached_cells_1;
        //    }
        //}
        public CKCells Cells => Cache(ref _cached_cells_1, () => RefreshCells());

        private CKCells RefreshCells()
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

                ckCells_1.Add(new CKCell(cellRef));
            }

            return new CKCells(ckCells_1, this);
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
            return null;
        }


        /// <summary>
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        [Obsolete("Planned for privatization")]

        public Word.Range COMRange
        {
            get
            {
                this.Ping();
                if (_COMRange == null || IsDirty)//isDirty?
                {

                    Refresh();
                    //throw new CKDebugException("Range is null");//just for tracking.
                }
                this.Pong();
                return _COMRange;
            }
            protected set
            {
                this.Ping();
                if (_COMRange != null) throw new CKDebugException("Attempted to assign a populated Range.");
                if (value is null) throw new ArgumentNullException("value");
                if (!Document.Matches(value)) throw new ArgumentException("value must share document refernce with host.");
                IsDirty = true;
                _COMRange = value;
                this.Pong();
            }
        }
        /// <summary>
        /// Unchecked COM range (typically call COMRange when the range might be dirty)
        /// </summary>
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

            get => Cache(ref _cachedText);
            set => SetCache(ref _cachedText, value, v => _COMRange.Text = v);
        }
        /// <summary>
        /// unsage.
        /// </summary>
        public Word.Font Font => _COMRange?.Font ?? throw new CKDebugException($"{nameof(_COMRange)} or Font was null.");

        /// <summary>
        /// Gets a "pretty" version of the range's text.
        /// This version replaces cell markers with tabs, preserves Windows-style newlines,
        /// and removes extraneous control characters.
        /// </summary>
        public string PrettyText => Cache(ref _cachedPrettyText);
        public bool IsCOMDeferred { get; private set; }


        protected T Cache<T>(ref T cachedField)
        {
            if (IsDirty || cachedField == null)
            {
                if (IsCOMDeferred)
                {
                    Log.Debug("Deferred COM access triggered inside Cache<T>.");
                    IsCOMDeferred = false;
                }
                Refresh();
            }
            return cachedField;
        }

        protected T Cache<T>(ref T cachedField, Func<T> refreshFunc)
        {
            if (IsDirty || cachedField == null)
            {
                cachedField = refreshFunc();
                if (IsCOMDeferred)
                {
                    Log.Debug("Deferred COM access triggered inside Cache<T> (custom refresh).");
                    IsCOMDeferred = false;
                }
            }
            return cachedField;
        }
        protected void SetCache<T>(ref T field, T value, Action<T> comSetter = null)
        {
            comSetter?.Invoke(value);
            field = value;
            IsDirty = true;
        }

        /// <summary>
        /// Gets the scrunched version of the range's text, i.e. all whitespace removed,
        /// for reliable comparisons.
        /// </summary>
        public string ScrunchedText => Cache(ref _cachedScrunchedText);


        /// <summary>
        /// Gets the starting position of the range.
        /// </summary>
        public int Start => Cache(ref _cachedStart, () => COMRange.Start);

        /// <summary>
        /// Gets the ending position of the range.
        /// </summary>
        public int End => Cache(ref _cachedEnd, () => COMRange.End);


        private bool _isCheckingDirty = false;
        private static long _isDirtyCount = 0;
        /// <inheritdoc/>
        public virtual bool IsDirty
        {
            get
            {
                this.Ping();
                if (_isDirtyCount++ % 20 == 0) LH.Checkpoint($"CKRange.IsDirty count: {_isDirtyCount}");

                if (_isCheckingDirty) return this.Pong(() => _isDirty);
                _isCheckingDirty = true;


                _isDirty = _isDirty || CheckDirtyFor();


                _isCheckingDirty = false;
                return this.Pong(() => _isDirty, msg: _isDirty.ToString());
            }
            set => _isDirty = value;
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
                return false;
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
        public CKTables Tables => new CKTables(COMRange.Tables, Document);
        /// <summary>
        /// Gets or sets the formatted text for this range. Setting this value replaces the contents and formatting of the range.
        /// </summary>
        public CKRange FormattedText
        {
            get
            {
                this.Ping();
                if (COMRange == null) throw new InvalidOperationException("COMRange is null.");
                var formatted = _COMRange.FormattedText;
                var result = new CKRange(formatted, Parent);
                this.Pong();
                return result;

            }
            set
            {
                this.Ping();
                if (COMRange == null) throw new InvalidOperationException("COMRange is null.");
                if (value?._COMRange == null) throw new ArgumentNullException(nameof(value));
                _COMRange.FormattedText = value._COMRange;
                IsDirty = true;
                this.Pong();

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
                collapsed = Document.Range(end, end)._COMRange;
            }
            catch
            {
                collapsed = docRange._COMRange; // fallback: entire document
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
                collapsed = Document.Range(start, start)._COMRange;
            }
            catch
            {
                collapsed = docRange._COMRange; // fallback: entire document TODO why?
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
            this.Ping();
            if (_isRefreshing) return;
            _isRefreshing = true;

            DoRefreshThings(); //sometimes COMRange could be assigned here.

            if (_COMRange is null) throw new InvalidOperationException($"{nameof(_COMRange)} cannot be null.");

            _cachedText = _COMRange.Text;
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


        #endregion

        private RangeSnapshot _snapshot = default;

        public RangeSnapshot Snapshot
        {
            get
            {
                this.Ping();
                if (_snapshot == null)
                {
                    _snapshot = new RangeSnapshot(_COMRange);
                }
                this.Pong();
                return _snapshot;
            }
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj == null || !(obj is CKRange other)) return false;
            return Equals(other);
        }

        public bool Equals(CKRange other)
        {
            if (ReferenceEquals(this, other)) return true;
            if (other == null) return false;

            // Use snapshot if available on both sides
            if (_snapshot != null && other._snapshot != null)
                return _snapshot.FastMatch(other._snapshot);

            // Fallback to slow COM-based snapshot comparison
            return Snapshot.SlowMatch(other._COMRange);
        }

        public override int GetHashCode()
        {
            // Prefer hash from snapshot if available
            if (_snapshot != null && !string.IsNullOrEmpty(_snapshot.FastHash))
                return _snapshot.FastHash.GetHashCode();

            // Fallback: use the COM object identity
            return Snapshot.FastHash.GetHashCode();
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
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                _disposedValue = true;
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
