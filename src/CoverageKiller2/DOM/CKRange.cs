using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a simple wrapper for the Word.Range object.
    /// </summary>
    public class CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        public Word.Range COMRange { get; private set; }

        /// <summary>
        /// Gets the starting position of the range.
        /// </summary>
        public int Start => COMRange.Start;

        public string Text => COMRange.Text;
        /// <summary>
        /// Gets the ending position of the range.
        /// </summary>
        public int End => COMRange.End;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is null.</exception>
        public CKRange(Word.Range range)
        {
            COMRange = range ?? throw new ArgumentNullException(nameof(range));
            _originalCharCount = COMRange.Characters.Count;
            _originalText = COMRange.Text;
        }

        private int _originalCharCount;
        private string _originalText;
        private bool _isDirty = false;

        public virtual bool IsDirty
        {
            get
            {
                _isDirty = _isDirty
                    || COMRange.Characters.Count != _originalCharCount
                    || COMRange.Text != _originalText;

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
            if (!(obj is CKRange))
                return false;
            CKRange other = (CKRange)obj;

            // Compare the underlying COMRange properties.
            return this.COMRange.Start == other.COMRange.Start &&
                   this.COMRange.End == other.COMRange.End;
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


    }
}
