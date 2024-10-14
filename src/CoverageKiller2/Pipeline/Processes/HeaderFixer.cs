using Serilog;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class HeaderFixer : CKWordPipelineProcess
    {
        private readonly CKDocument template;

        public HeaderFixer(CKDocument template)
        {
            this.template = template;
        }

        public override void Process()
        {
            Log.Information("Fixing Header");
            var wordApp = CKDoc.WordApp;
            WordSelector.Header(template);
            wordApp.Selection.Copy();

            CKDoc.Activate();
            wordApp.Selection.GoTo(
                What: Word.WdGoToItem.wdGoToSection,
                Which: Word.WdGoToDirection.wdGoToAbsolute,
                Count: 1);

            WordSelector.Header(CKDoc);
            wordApp.Selection.PasteAndFormat(
                Word.WdRecoveryType.wdFormatOriginalFormatting);

            Clipboard.Clear();

        }


    }
}