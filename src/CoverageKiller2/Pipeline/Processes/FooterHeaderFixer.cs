//using Microsoft.Office.Interop.Word;
using CoverageKiller2.DOM;
using Serilog;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class FooterHeaderFixer : CKWordPipelineProcess
    {
        public FooterHeaderFixer(CKDocument template)
        {
            Template = template;
        }
        public FooterHeaderFixer()
        {

        }
        public override void Process()
        {
            Log.Information("Fixing Footer.");
            Template.CopyHeaderAndFooterTo(CKDoc);
            //var wordApp = CKDoc.WordApp;
            //template.SelectFooterWholeStory(); //needs to be moved out to CKDocument, not an extension to Word.Document.
            //wordApp.Selection.Copy();

            //CKDoc.Activate();
            //wordApp.Selection.GoTo(
            //    What: Word.WdGoToItem.wdGoToSection,
            //    Which: Word.WdGoToDirection.wdGoToAbsolute,
            //    Count: 1);

            //CKDoc.COMObject.SelectFooterWholeStory();
            //wordApp.Selection.Delete();
            //wordApp.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

            //Clipboard.Clear();
        }

    }
}