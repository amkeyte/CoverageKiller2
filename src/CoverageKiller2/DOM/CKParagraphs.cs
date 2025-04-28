using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of <see cref="CKParagraph"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.03.0000
    /// </remarks>
    public class CKParagraphs : ACKRangeCollection, IEnumerable<CKParagraph>
    {
        private List<CKParagraph> _cachedParagraphs;

        /// <summary>
        /// Gets the underlying Word.Paragraphs COM object from the parent range.
        /// May be null if deferred.
        /// </summary>
        protected Word.Paragraphs COMParagraphs => Cache(ref _COMParagraphs, () =>
        {
            if (Parent is CKRange parentRange)
            {
                // Someday CKRange will lazily cache its COMRange better too
                return parentRange.COMRange.Paragraphs;
            }
            throw new CKDebugException("Unsupported parent type for resolving COMParagraphs.");
        });

        private Word.Paragraphs _COMParagraphs = default;
        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class.
        /// </summary>
        /// <param name="collection">The Word.Paragraphs COM collection to wrap.</param>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        /// <param name="deferCOM">If true, initializes paragraphs in defer mode.</param>
        public CKParagraphs(Word.Paragraphs collection, IDOMObject parent, bool deferCOM = false)
            : base(parent, deferCOM)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (collection.Count > 1000)
                throw new ArgumentException($"Paragraphs collection is too large ({collection.Count}). Find a smaller subset.");

            Log.Information($"CKParagraphs created with {collection.Count} entries (deferCOM={deferCOM}).");

            COMParagraphs = collection;
            InitializeEmptyList();
        }

        public CKParagraphs(IDOMObject parent, bool deferCOM = true)
             : base(parent, deferCOM)
        {

        }

        /// <inheritdoc/>
        public override int Count => COMParagraphs?.Count ?? 0;

        /// <inheritdoc/>
        public override bool IsOrphan => Parent.IsOrphan;

        private void InitializeEmptyList()
        {
            _cachedParagraphs = new List<CKParagraph>(COMParagraphs?.Count ?? 0);
            for (int i = 0; i < (COMParagraphs?.Count ?? 0); i++)
            {
                _cachedParagraphs.Add(null); // placeholder nulls
            }
        }

        /// <summary>
        /// Gets the <see cref="CKParagraph"/> at the specified one-based index.
        /// </summary>
        public CKParagraph this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of paragraphs.");

                var para = _cachedParagraphs[index - 1];
                if (para == null)
                {
                    if (IsCOMDeferred)
                    {
                        para = new CKParagraph(this, index);
                    }
                    else
                    {
                        para = new CKParagraph(COMParagraphs[index], this);
                    }
                    _cachedParagraphs[index - 1] = para;
                }
                return para;
            }
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            if (obj is CKRange para)
            {
                for (int i = 0; i < _cachedParagraphs.Count; i++)
                {
                    if (_cachedParagraphs[i] != null && _cachedParagraphs[i].Equals(para))
                        return i + 1; // one-based
                }
            }
            return -1;
        }

        /// <inheritdoc/>
        public IEnumerator<CKParagraph> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <inheritdoc/>
        public override void Clear()
        {
            _cachedParagraphs?.Clear();
            _cachedParagraphs = null;
            _isDirty = true;
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"CKParagraphs [Count: {Count}]";
        }

        /// <inheritdoc/>
        protected override void DoRefreshThings()
        {
            InitializeEmptyList();
            _isDirty = false;
        }
    }
}
