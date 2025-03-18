using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents a wrapper for the Word.Section object, exposing common functionality for a section in a Word document.
    /// </summary>
    public class CKSection
    {
        /// <summary>
        /// Gets the underlying Word.Section COM object.
        /// </summary>
        internal Word.Section COMObject { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKSection"/> class.
        /// </summary>
        /// <param name="section">The Word.Section object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the section parameter is null.</exception>
        public CKSection(Word.Section section)
        {
            COMObject = section ?? throw new ArgumentNullException(nameof(section));
        }

        /// <summary>
        /// Gets the range of the section.
        /// </summary>
        public CKRange Range => CKRange.Create(COMObject.Range);

        /// <summary>
        /// Gets or sets the primary header range for this section.
        /// Setting this property replaces the formatted content of the header with that of the provided range.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when the value is null.</exception>
        public Word.Range HeaderRange
        {
            get => COMObject.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                COMObject.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.FormattedText = value.FormattedText;
            }
        }

        /// <summary>
        /// Gets or sets the primary footer range for this section.
        /// Setting this property replaces the formatted content of the footer with that of the provided range.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when the value is null.</exception>
        public Word.Range FooterRange
        {
            get => COMObject.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                COMObject.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.FormattedText = value.FormattedText;
            }
        }

        /// <summary>
        /// Gets the page setup properties for this section.
        /// </summary>
        public Word.PageSetup PageSetup => COMObject.PageSetup;

        /// <summary>
        /// Returns a string representation of the current section.
        /// </summary>
        /// <returns>A string representing the section's range.</returns>
        public override string ToString()
        {
            return $"Section: Range [{Range.Start}, {Range.End}]";
        }

        /// <summary>
        /// Gets an enumerable collection of <see cref="CKTable"/> objects contained within the section.
        /// </summary>
        public CKTables Tables => CKTables.Create(Range);

    }
}
