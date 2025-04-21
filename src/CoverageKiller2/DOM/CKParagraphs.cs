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
    /// Version: CK2.00.01.0000
    /// </remarks>
    public class CKParagraphs : ACKRangeCollection, IEnumerable<CKParagraph>
    {
        private List<CKParagraph> _cachedParagraphs;

        /// <summary>
        /// Gets the underlying Word.Paragraphs COM object from the parent range.
        /// </summary>
        public Word.Paragraphs COMParagraphs { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        /// <exception cref="ArgumentNullException">Thrown when the parent parameter is null.</exception>
        public CKParagraphs(Word.Paragraphs collection, IDOMObject parent) : base(parent)
        {
            Log.Information($"CKParagraps was created with {collection.Count} entries.");

            if (collection.Count > 1000) throw new ArgumentException($"Paragraphs collection is to large({collection.Count}). Find a smaller subset.");
            COMParagraphs = collection;

        }

        /// <summary>
        /// Gets the number of paragraphs in the associated range.
        /// </summary>
        public override int Count => ParagraphList.Count;

        /// <summary>
        /// Gets whether the paragraph cache is dirty.
        /// </summary>
        public override bool IsDirty { get; protected set; }

        /// <summary>
        /// Gets whether the paragraph collection is orphaned.
        /// </summary>
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
        /// Gets the internal list of cached paragraphs, refreshing if dirty.
        /// </summary>
        private List<CKParagraph> ParagraphList
        {
            get
            {
                if (_cachedParagraphs == null || IsDirty)
                {
                    _cachedParagraphs = new List<CKParagraph>();
                    for (int i = 1; i <= COMParagraphs.Count; i++)
                    {
                        _cachedParagraphs.Add(new CKParagraph(COMParagraphs[i], this));
                    }
                    IsDirty = false;
                }
                return _cachedParagraphs;
            }
        }

        /// <summary>
        /// Returns the one-based index of the specified object, or -1 if not found.
        /// </summary>
        /// <param name="obj">The object to locate in the collection.</param>
        /// <returns>The one-based index of the object, or -1 if not found.</returns>
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

        /// <summary>
        /// Note: Iterating Pagraphs can be VERY slow for large collections. Fix this somehow.
        /// Returns an enumerator that iterates through the collection of <see cref="CKParagraph"/> objects.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKParagraph"/> objects.</returns>
        public IEnumerator<CKParagraph> GetEnumerator()
        {
            return ParagraphList.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator for the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// Returns a string representation of the CKParagraphs collection.
        /// </summary>
        /// <returns>A string containing the count of paragraphs.</returns>
        public override string ToString()
        {
            return $"CKParagraphs [Count: {Count}]";
        }

        public override void Clear()
        {
            _cachedParagraphs.Clear();
        }
    }
}
