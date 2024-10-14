//using Microsoft.Office.Interop.Word;
using Serilog;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class FooterFixer : CKWordPipelineProcess
    {
        private IndoorReportTemplate template;

        public FooterFixer(IndoorReportTemplate template)
        {
            this.template = template;
        }

        public override void Process()
        {
            Log.Information("Fixing Footer.");
            var wordApp = CKDoc.WordApp;
            template.WordDoc.SelectFooterWholeStory(); //needs to be moved out to CKDocument, not an extension to Word.Document.
            wordApp.Selection.Copy();

            CKDoc.Activate();
            wordApp.Selection.GoTo(
                What: Word.WdGoToItem.wdGoToSection,
                Which: Word.WdGoToDirection.wdGoToAbsolute,
                Count: 1);

            CKDoc.WordDoc.SelectFooterWholeStory();
            wordApp.Selection.Delete();
            wordApp.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

            Clipboard.Clear();
        }
    }
}