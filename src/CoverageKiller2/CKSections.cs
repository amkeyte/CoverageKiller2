using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents a collection of <see cref="CKSection"/> objects associated with a <see cref="CKDocument"/>.
    /// </summary>
    public class CKSections : IEnumerable<CKSection>
    {

        internal static CKSections Create(CKDocument parent)
        {
            parent = parent ?? throw new ArgumentNullException(nameof(parent));
            return new CKSections(parent);
        }



        /// <summary>
        /// Gets the parent <see cref="CKDocument"/> that contains this collection of sections.
        /// </summary>
        public CKDocument Parent { get; private set; }

        /// <summary>
        /// Gets the underlying Word.Sections COM object from the parent document.
        /// </summary>
        internal Word.Sections COMObject => Parent.COMObject.Sections;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKSections"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKDocument"/>.</param>
        /// <exception cref="ArgumentNullException">Thrown when the parent parameter is null.</exception>
        internal CKSections(CKDocument parent)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        /// <summary>
        /// Gets the number of sections in the document.
        /// </summary>
        public int Count => COMObject.Count;

        /// <summary>
        /// Gets the <see cref="CKSection"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the section to retrieve.</param>
        /// <returns>The <see cref="CKSection"/> at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the index is less than 1 or greater than the number of sections.
        /// </exception>
        public CKSection this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of sections.");
                return new CKSection(COMObject[index]);
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection of <see cref="CKSection"/> objects.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKSection"/> objects.</returns>
        public IEnumerator<CKSection> GetEnumerator()
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
        /// Returns a string representation of the CKSections collection.
        /// </summary>
        /// <returns>A string containing the document name and the number of sections.</returns>
        public override string ToString()
        {
            string docName = System.IO.Path.GetFileName(Parent.FullPath);
            return $"[{docName}].Sections[Count: {Count}]";
        }
    }
}
