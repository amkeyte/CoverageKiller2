using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a basic wrapper for a Word.Paragraph object.
    /// Inherits from CKRange since a paragraph is fundamentally a range.
    /// </summary>
    public class CKParagraph : CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Paragraph COM object.
        /// May be null if deferred and not yet realized.
        /// </summary>
        public Word.Paragraph COMParagraph { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraph"/> class
        /// with an immediate Word.Paragraph object.
        /// </summary>
        /// <param name="paragraph">The Word.Paragraph object to wrap.</param>
        /// <param name="parent">The parent DOM object.</param>
        /// <exception cref="ArgumentNullException">Thrown when the paragraph parameter is null.</exception>
        public CKParagraph(Word.Paragraph paragraph, IDOMObject parent)
            : base(paragraph?.Range, parent)
        {
            COMParagraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraph"/> class
        /// in deferred mode without an immediate Word.Paragraph.
        /// </summary>
        /// <param name="parent">The parent DOM object.</param>
        public CKParagraph(IDOMObject parent)
            : base(parent) // triggers deferCOM = true
        {
            COMParagraph = null;
        }

        /// <summary>
        /// Returns a string representation of the CKParagraph.
        /// </summary>
        /// <returns>A string showing the paragraph's range.</returns>
        public override string ToString()
        {
            return $"CKParagraph: Range [{Start}, {End}]";
        }

        /// <summary>
        /// Static constructor to register casting logic.
        /// </summary>
        static CKParagraph()
        {
            IDOMCaster.Register(input =>
            {
                CKParagraph result = default;

                if (input.Parent is CKDocument doc)
                {
                    var index = doc.Range().Paragraphs.IndexOf(input);
                    result = doc.Range().Paragraphs[index];
                }
                else if (input.Parent is CKParagraphs paras)
                {
                    result = paras[paras.IndexOf(input)];
                }
                else if (input.Parent is CKRange rng)
                {
                    result = rng.Paragraphs[1];
                }

                return result ?? throw new InvalidCastException("Could not convert to CKParagraph.");
            });
        }
    }
}
