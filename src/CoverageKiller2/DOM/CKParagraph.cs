using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a basic wrapper for the Word.Paragraph object.
    /// Inherits from CKRange since a paragraph is fundamentally a range.
    /// </summary>
    public class CKParagraph : CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Paragraph COM object.
        /// </summary>
        public Word.Paragraph COMParagraph { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraph"/> class.
        /// </summary>
        /// <param name="paragraph">The Word.Paragraph object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the paragraph parameter is null.</exception>
        public CKParagraph(Word.Paragraph paragraph) : base(paragraph.Range)
        {
            COMParagraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
        }

        /// <summary>
        /// Returns a string representation of the CKParagraph.
        /// </summary>
        /// <returns>A string showing the paragraph's range.</returns>
        public override string ToString()
        {
            return $"CKParagraph: Range [{Start}, {End}]";
        }
    }
}
