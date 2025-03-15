using Serilog;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class DefaultFixer2 : CKWordPipelineProcess
    {
        private readonly CKDocument template;

        public DefaultFixer2(CKDocument template)
        {
            this.template = template;
        }

        public override void Process()
        {
            Log.Information("DefaultFixer1 Processing");
        }
    }
}
