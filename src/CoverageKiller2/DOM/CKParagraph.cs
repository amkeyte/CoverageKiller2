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

        private readonly int _paragraphIndex; // 1-based

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraph"/> class with an immediate Word.Paragraph object.
        /// </summary>
        /// <param name="paragraph">The Word.Paragraph object to wrap.</param>
        /// <param name="parent">The parent DOM object.</param>
        public CKParagraph(Word.Paragraph paragraph, IDOMObject parent)
            : base(paragraph?.Range, parent)
        {
            COMParagraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            _paragraphIndex = -1; // not needed if we already have the paragraph
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKParagraph"/> class in deferred mode with a known paragraph index.
        /// </summary>
        /// <param name="parent">The parent DOM object (must be a CKParagraphs).</param>
        /// <param name="index">The 1-based paragraph index within the parent Word.Paragraphs collection.</param>
        public CKParagraph(IDOMObject parent, int index)
            : base(parent) // deferCOM = true
        {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
            _paragraphIndex = index;
            COMParagraph = null;
        }

        /// <summary>
        /// Ensures that the COMParagraph reference is available, resolving it if necessary.
        /// </summary>
        private void EnsureCOMParagraphReady()
        {
            if (COMParagraph == null)
            {
                if (Parent is CKParagraphs paras)
                {
                    if (paras.COMParagraphs == null)
                        throw new InvalidOperationException("Deferred CKParagraph has no valid COMParagraphs collection.");

                    if (_paragraphIndex < 1 || _paragraphIndex > paras.COMParagraphs.Count)
                        throw new InvalidOperationException("Deferred CKParagraph has invalid index.");

                    COMParagraph = paras.COMParagraphs[_paragraphIndex];
                    _COMRange = COMParagraph.Range;
                    _deferCOM = false;
                    Refresh();
                }
                else
                {
                    throw new InvalidOperationException("Deferred CKParagraph must have CKParagraphs as parent.");
                }
            }
        }

        /// <inheritdoc/>
        public override string Text
        {
            get
            {
                EnsureCOMParagraphReady();
                return base.Text;
            }
        }

        /// <inheritdoc/>
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
