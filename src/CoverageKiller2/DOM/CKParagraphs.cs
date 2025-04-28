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
    /// Version: CK2.00.02.0000
    /// </remarks>
    public class CKParagraphs : ACKRangeCollection, IEnumerable<CKParagraph>
    {
        private List<CKParagraph> _cachedParagraphs;

        /// <summary>
        /// Gets the underlying Word.Paragraphs COM object from the parent range.
        /// May be null if deferred.
        /// </summary>
        public Word.Paragraphs COMParagraphs { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class.
        /// </summary>
        /// <param name="collection">The Word.Paragraphs COM collection to wrap.</param>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        /// <param name="deferCOM">If true, initializes paragraphs in defer mode.</param>
        /// <exception cref="ArgumentNullException">Thrown when the parent parameter is null.</exception>
        public CKParagraphs(Word.Paragraphs collection, IDOMObject parent, bool deferCOM = false)
            : base(parent, deferCOM)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (collection.Count > 1000)
                throw new ArgumentException($"Paragraphs collection is too large ({collection.Count}). Find a smaller subset.");

            Log.Information($"CKParagraphs created with {collection.Count} entries (deferCOM={deferCOM}).");

            COMParagraphs = collection;
        }

        /// <inheritdoc/>
        public override int Count => ParagraphList.Count;

        /// <inheritdoc/>
        public override bool IsOrphan => Parent.IsOrphan;

        /// <summary>
        /// Gets the <see cref="CKParagraph"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the paragraph to retrieve.</param>
        /// <returns>The <see cref="CKParagraph"/> at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the index is less than 1 or greater than the number of paragraphs.
        /// </exception>
        public CKParagraph this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of paragraphs.");
                return ParagraphList[index - 1];
            }
        }

        /// <summary>
        /// Gets the internal list of paragraphs, refreshing if dirty.
        /// </summary>
        private List<CKParagraph> ParagraphList
        {
            get
            {
                return Cache(ref _cachedParagraphs, BuildParagraphList);
            }
        }

        /// <summary>
        /// Builds the internal paragraph list.
        /// This does not touch COMParagraphs if defer is still active.
        /// </summary>
        private List<CKParagraph> BuildParagraphList()
        {
            var list = new List<CKParagraph>();

            if (_deferCOM)
            {
                // Defer: create empty placeholder paragraphs
                Log.Debug("Building deferred CKParagraphs.");
                for (int i = 1; i <= COMParagraphs?.Count; i++)
                {
                    list.Add(new CKParagraph(this)); // Deferred CKParagraphs
                }
            }
            else
            {
                // Immediate: wrap real COM paragraphs
                Log.Debug("Building full CKParagraphs with COM access.");
                for (int i = 1; i <= COMParagraphs.Count; i++)
                {
                    list.Add(new CKParagraph(COMParagraphs[i], this));
                }
            }

            return list;
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            if (obj is CKRange para)
            {
                for (int i = 0; i < ParagraphList.Count; i++)
                {
                    if (ParagraphList[i].Equals(para))
                        return i + 1; // one-based
                }
            }
            return -1;
        }

        /// <inheritdoc/>
        public IEnumerator<CKParagraph> GetEnumerator()
        {
            return ParagraphList.GetEnumerator();
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

        /// <summary>
        /// Forces a refresh of the paragraph list.
        /// </summary>
        protected override void Refresh()
        {
            _cachedParagraphs = BuildParagraphList();
            _isDirty = false;
        }
    }
}
