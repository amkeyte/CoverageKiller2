using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a simple wrapper for the Word.Range object.
    /// </summary>
    public class CKRange
    {

        private string _cachedText;
        private string _cachedPrettyText;
        private string _cachedScrunchedText;

        /// <summary>
        /// Updates cached text values from the underlying COMRange and resets the dirty flag.
        /// </summary>
        public void Refresh()
        {
            _cachedText = COMRange.Text;
            // Reset original values for future dirty checks.
            //_originalText = COMRange.Text;
            _cachedCharCount = COMRange.Characters.Count;

            _cachedPrettyText = CKTextHelper.Pretty(_cachedText);
            _cachedScrunchedText = CKTextHelper.Scrunch(_cachedText);
            IsDirty = false;
        }

        /// <summary>
        /// Gets the raw text of the range.
        /// </summary>
        public string Text
        {
            get
            {
                if (IsDirty || _cachedText == null)
                    Refresh();
                return _cachedText;
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
        /// Gets the scrunched version of the range's text,
        /// i.e. all whitespace removed, for reliable comparisons.
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
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        public Word.Range COMRange { get; private set; }

        /// <summary>
        /// Gets the starting position of the range.
        /// </summary>
        public int Start => COMRange.Start;

        /// <summary>
        /// Gets the ending position of the range.
        /// </summary>
        public int End => COMRange.End;

        /// <summary>
        /// Gets the raw text of the range as returned by Word.
        /// </summary>
        public string RawText => COMRange.Text;

        /// <summary>
        /// Gets a "pretty" version of the range's text.
        /// This version replaces cell markers with tabs, preserves Windows-style newlines,
        /// and removes extraneous control characters.
        /// </summary>


        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the range parameter is null.</exception>
        public CKRange(Word.Range range)
        {
            COMRange = range ?? throw new ArgumentNullException(nameof(range));
            _cachedCharCount = COMRange.Characters.Count;
            _cachedText = COMRange.Text;
        }

        private int _cachedCharCount;
        //private string _originalText;
        private bool _isDirty = false;

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

        public CKSections Sections => new CKSections(this);
        public CKParagraphs Paragraphs => new CKParagraphs(this);
        public CKTables Tables => new CKTables(this);
        public CKCells Cells => new CKCells(this);

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;
            if (obj == null)
                return false;
            if (!(obj is CKRange))
                return false;
            return Equals((CKRange)obj);
        }

        public bool Equals(CKRange other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other == null)
                return false;

            // Compare the underlying COMRange boundaries.
            return this.COMRange.Start == other.COMRange.Start &&
                   this.COMRange.End == other.COMRange.End;
        }

        public bool TextEquals(string other)
        {
            // Get the scrunched version of this range's text.
            string Scrunched = CKTextHelper.Scrunch(this.COMRange.Text);
            // Get the scrunched version of the provided text.
            string otherScrunched = CKTextHelper.Scrunch(other);
            // Compare the normalized strings using ordinal comparison.
            return string.Equals(Scrunched, otherScrunched, StringComparison.Ordinal);
        }
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + COMRange.Start.GetHashCode();
                hash = hash * 23 + COMRange.End.GetHashCode();
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
    }
}
