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
        public CKParagraph(Word.Paragraph paragraph, IDOMObject parent) : base(paragraph?.Range, parent)
        {
            //LH.Ping(GetType());
            COMParagraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            //LH.Pong(GetType());
        }

        /// <summary>
        /// Returns a string representation of the CKParagraph.
        /// </summary>
        /// <returns>A string showing the paragraph's range.</returns>
        public override string ToString()
        {
            return $"CKParagraph: Range [{Start}, {End}]";
        }

        static CKParagraph()
        {
            IDOMCaster.Register(input =>
            {

                CKParagraph result = default;

                if (input.Parent is CKDocument doc)
                {
                    //do the doc thing.
                    var index = doc.Range().Paragraphs.IndexOf(input);
                    result = doc.Range().Paragraphs[index];
                }
                else if (input.Parent is CKParagraphs paras)
                {
                    //do the paragrapgs is parent thing
                    result = paras[paras.IndexOf(input)];
                }
                else if (input.Parent is CKRange rng)
                {
                    result = rng.Paragraphs[1];
                    //do the range thing
                }

                return result ?? throw new InvalidCastException("Cound not convert to CKParagraph.");
            });
        }


    }
}
