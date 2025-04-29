using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a wrapper for the Word.Section object, exposing common functionality for a section in a Word document.
    /// </summary>
    public class CKSection : CKRange
    {
        /// <summary>
        /// Do not use: probably will be hidden.
        /// Gets the underlying Word.Section COM object.
        /// </summary>
        public Word.Section COMSection { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKSection"/> class.
        /// </summary>
        /// <param name="section">The Word.Section object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the section parameter is null.</exception>
        public CKSection(Word.Section section, IDOMObject parent) : base(section?.Range, parent)
        {
            COMSection = (Word.Section)section ?? throw new ArgumentNullException(nameof(section));
        }

        /// <summary>
        /// Gets or sets the primary header range for this section.
        /// Setting this property replaces the formatted content of the header with that of the provided range.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when the value is null.</exception>
        public CKRange HeaderRange
        {
            get => new CKRange(COMSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, this);
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                COMSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.FormattedText = value.COMRange.FormattedText;
            }
        }

        /// <summary>
        /// Gets or sets the primary footer range for this section.
        /// Setting this property replaces the formatted content of the footer with that of the provided range.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when the value is null.</exception>
        public CKRange FooterRange
        {
            get => new CKRange(COMSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, this);
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                COMSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.FormattedText = value.COMRange.FormattedText;
            }
        }

        /// <summary>
        /// Gets the page setup properties for this section.
        /// </summary>
        public Word.PageSetup PageSetup => COMSection.PageSetup;

        /// <summary>
        /// Returns a string representation of the current section.
        /// </summary>
        /// <returns>A string representing the section's range.</returns>
        public override string ToString()
        {
            return $"Section: Range [{Start}, {End}]";
        }



    }
}
