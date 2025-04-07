using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;


///**********
///When you access the Sections property on a Word.Range, Word returns 
///a collection of the sections that intersect with that range. This means:
//
//Contained in a Single Section:
//If your range is entirely within one section (even if it doesn't cover
//the whole section), it will return a collection with that one section.

//Spanning Multiple Sections:
//If the range spans across a section break, the returned collection will include
//each section that the range touches, even if only partially.

//No Failure:
//The property doesn't fail or throw an error if the range is
//smaller than a section. It simply returns the sections that intersect
//the range—even if that means just one section (or, in unusual cases, none
//if the range is somehow empty).

//In summary, VSTO (via the Word Interop) gracefully returns the intersecting
//sections without error, making it safe to use even when your range doesn't
//encompass an entire section.
///**********
///**********
//Range is “live” and will update to show new sections that fall within the Range’s current boundaries. 
//
//In summary, like Document.Sections, Range.Sections is a live collection that updates as
//long as the change occurs within the range’s defined boundaries.

///**********




namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of <see cref="CKSection"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    public class CKSections : ACKRangeCollection, IEnumerable<CKSection>
    {
        /// <summary>
        /// Probably will get hidden. Avoid use if possible.
        /// </summary>
        public Word.Sections COMSection => Parent.COMRange.Sections;


        /// <summary>
        /// Initializes a new instance of the <see cref="CKSections"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        /// <exception cref="ArgumentNullException">Thrown when the parent parameter is null.</exception>
        public CKSections(CKRange parent) : base(parent)
        {
        }

        /// <summary>
        /// Gets the number of sections in the range.
        /// </summary>
        public override int Count => COMSection.Count;

        public override bool IsDirty => throw new NotImplementedException();

        public override bool IsOrphan => throw new NotImplementedException();

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
                return new CKSection(COMSection[index]);
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
        /// <returns>A string containing the count of sections.</returns>
        public override string ToString()
        {
            return $"CKSections [Count: {Count}]";
        }
    }
}
