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
    public class CKRange : IDOMObject
    {
        #region Fields & Caching

        private string _cachedText;
        private string _cachedPrettyText;
        private string _cachedScrunchedText;
        private int _cachedCharCount;
        private int _cachedStart;
        private int _cachedEnd;
        private bool _isDirty = false;
        private bool _isOrphan = false;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <param name="parent">Optional parent DOM object; if not provided, will be looked up via CKDocuments.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> parameter is null.</exception>
        public CKRange(Word.Range range, IDOMObject parent = null)
        {
            COMRange = range ?? throw new ArgumentNullException(nameof(range));
            Parent = parent ?? CKDocuments.GetByCOMDocument(COMRange.Document);
            _cachedCharCount = COMRange.Characters.Count;
            _cachedText = COMRange.Text;
            // Initialize cached boundary values.
            _cachedStart = COMRange.Start;
            _cachedEnd = COMRange.End;
        }

        #endregion

        #region Public Properties

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
                    return false;
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.Message);
                    _isOrphan = true;
                    return true;
                }
                catch (Exception)
                {
                    _isOrphan = true;
                    return true;
                }
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
        public Word.Application Application => Document.Application;

        /// <summary>
        /// Gets the sections contained in the range.
        /// </summary>
        public CKSections Sections => new CKSections(this);

        /// <summary>
        /// Gets the paragraphs contained in the range.
        /// </summary>
        public CKParagraphs Paragraphs => new CKParagraphs(this);

        /// <summary>
        /// Gets the tables contained in the range.
        /// </summary>
        public CKTables Tables => new CKTables(this);

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

        #region Equality Overrides

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

            int thisStart = IsOrphan ? _cachedStart : COMRange.Start;
            int thisEnd = IsOrphan ? _cachedEnd : COMRange.End;
            int otherStart = other.IsOrphan ? other._cachedStart : other.COMRange.Start;
            int otherEnd = other.IsOrphan ? other._cachedEnd : other.COMRange.End;

            return thisStart == otherStart && thisEnd == otherEnd;
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

        #endregion

    }
    public static class RangeExtensions
    {
        public static bool Contains(this Word.Range outer, Word.Range inner)
        {
            return inner.Start >= outer.Start && inner.End <= outer.End;
        }
    }
}
