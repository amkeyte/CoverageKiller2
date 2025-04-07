using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of <see cref="CKParagraph"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    public class CKParagraphs : ACKRangeCollection, IEnumerable<CKParagraph>
    {
        /// <summary>
        /// Gets the underlying Word.Paragraphs COM object from the parent range.
        /// </summary>
        public Word.Paragraphs COMParagraphs => Parent.COMRange.Paragraphs;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraphs"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        /// <exception cref="ArgumentNullException">Thrown when the parent parameter is null.</exception>
        public CKParagraphs(CKRange parent) : base(parent) { }

        /// <summary>
        /// Gets the number of paragraphs in the associated range.
        /// </summary>
        public override int Count => COMParagraphs.Count;

        public override bool IsDirty => throw new NotImplementedException();

        public override bool IsOrphan => throw new NotImplementedException();

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
                return new CKParagraph(COMParagraphs[index]);
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection of <see cref="CKParagraph"/> objects.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKParagraph"/> objects.</returns>
        public IEnumerator<CKParagraph> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
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
    }
}
