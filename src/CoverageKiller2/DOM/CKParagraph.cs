using CoverageKiller2.Logging;
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
        protected Word.Paragraph COMParagraph
        {
            get => Cache(ref _COMParagraph, () =>
            {
                //here, either IsDirty is true or COMParagraph is null. Deferred is either. Refresh must be called.



                if (Parent is CKParagraphs parent && parent.Parent is CKRange parentRange)
                {
                    var comParas = parentRange.COMRange.Paragraphs;

                    if (Index < 1 || Index > comParas.Count)
                        throw new CKDebugException($"Invalid index {Index} for given Paragraphs collection.");

                    var comPara = comParas[Index];

                    if (COMRange is null) COMRange = comPara.Range; // do not use ?? here

                    return comPara;
                }
                throw new InvalidOperationException("Unsupported parent type for resolving COMParagraphs.");

            });

            private set => SetCache(ref _COMParagraph, value, (v) => COMParagraph = v);
        }


        private Word.Paragraph _COMParagraph;

        public CKParagraph(Word.Paragraph paragraph, IDOMObject parent, int index = -1)
            : base(paragraph?.Range, parent)
        {
            COMParagraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));

            Index = index;
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
            Index = index;
            COMParagraph = null;
        }

        protected override void DoRefreshThings()
        {
            this.Ping();
            if (IsCOMDeferred)
            {

            }
            else
            {

                //checked if it's null to force COMParagraph to update, so that COMRange is valid.
                if (COMParagraph == null) throw new CKDebugException("COMParagraph cannot refresh.");
            }
            this.Pong();
        }

        public int Index { get; private set; }

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
