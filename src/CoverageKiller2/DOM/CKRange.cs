using CoverageKiller2.Logging;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a simple wrapper for the Word.Range object.
    /// </summary>
    public class CKRange : IDOMObject
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
            _cachedCharCount = COMRange.Characters.Count;
            _cachedStart = COMRange.Start;
            _cachedEnd = COMRange.End;
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
        public CKRange(Word.Range range, IDOMObject parent = null)
        {
            LH.Ping(GetType());
            COMRange = range ?? throw new ArgumentNullException(nameof(range));
            Parent = parent ?? CKDocuments.GetByCOMDocument(COMRange.Document);
            _cachedCharCount = COMRange.Characters.Count;
            _cachedText = COMRange.Text;
            LH.Pong(GetType());

        }



        private int _cachedCharCount;
        private int _cachedStart;
        private int _cachedEnd;

        //private string _originalText;
        private bool _isDirty = false;
        private bool _isOrphan = false;

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
        /// Gets a value indicating whether this CKDocument no longer has a valid COMDocument reference.
        /// This becomes true if the document is closed or the COM object has been released.
        /// </summary>
        public bool IsOrphan
        {
            get
            {
                if (_isOrphan) return true;
                try
                {
                    // Accessing COMDocument.Application should throw if the COM object is no longer valid.
                    // Alternatively, accessing COMDocument.FullName is often sufficient.
                    _ = COMRange.Text;
                    return false;
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.Message);
                    _isOrphan = true;
                    return true;
                }
                catch (Exception ex)
                {
                    _isOrphan = true;
                    return true;
                }
            }
        }

        public CKSections Sections => new CKSections(this);
        public CKParagraphs Paragraphs => new CKParagraphs(this);
        public CKTables Tables => new CKTables(this);
        public CKCells Cells => new CKCells(this);

        public CKDocument Document => Parent.Document;

        public Word.Application Application => Document.Application;

        public IDOMObject Parent { get; private set; }

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

            int thisStart = IsOrphan ? _cachedStart : COMRange.Start;
            int thisEnd = IsOrphan ? _cachedEnd : COMRange.End;
            int otherStart = other.IsOrphan ? other._cachedStart : other.COMRange.Start;
            int otherEnd = other.IsOrphan ? other._cachedEnd : other.COMRange.End;

            return thisStart == otherStart && thisEnd == otherEnd;
        }

        public bool TextEquals(string other)
        {
            return CKTextHelper.ScrunchEquals(Text, other);
        }
        public override int GetHashCode()
        {
            //Summary of RPC Problem Resolution in CKRange and CKTableGrid
            //
            //We encountered RPC errors when accessing CKRange objects that became orphaned
            //(i.e., their underlying COM objects were no longer valid because the document
            //had been closed). The root cause was that our CKRange.Equals and GetHashCode
            //methods directly accessed live COM properties(COMRange.Start and COMRange.End),
            //which triggered exceptions when those COM objects were released.

            //The patch included two key changes:

            //1 Caching in CKRange:
            //  We updated the CKRange.Refresh() method to cache the current COMRange
            //  values(including Start and End) in private fields.Then, in Equals and GetHashCode,
            //  we check if the range is orphaned(using the IsOrphan property). If so, we fall back to
            //  the cached values instead of accessing the live COM object. This ensures that orphaned
            //  CKRange objects can still be compared and used in dictionaries without throwing RPC errors.
            //
            //2 Purging the Static Dictionary in CKTableGrid:
            //  We added logic to purge stale CKRange keys from the static dictionary of CKTableGrid
            //  before using it. This prevents orphaned references from lingering in the dictionary,
            //  ensuring that lookups use only valid, up-to-date ranges.
            //
            //These changes effectively resolve the RPC error by decoupling our equality and hash
            //  code computations from the live COM objects, thereby preserving functionality (such
            //  as dictionary lookups) even after a document is closed without shutting down the Word
            //  application.
            //
            //
            //This fix is not robust and should be given more attention later.

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
    }
}
