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

        public override void Process()
        {
            Log.Information("Fixing Header / Footer.");
            Template.CopyHeaderTo(CKDoc);
            Template.CopyFooterTo(CKDoc);
        }

    }
}