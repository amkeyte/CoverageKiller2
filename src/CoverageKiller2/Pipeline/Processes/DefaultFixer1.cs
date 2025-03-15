using Serilog;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class DefaultFixer1 : CKWordPipelineProcess
    {
        private readonly CKDocument template;

        public DefaultFixer1(CKDocument template)
        {
            this.template = template;
        }

        public override void Process()
        {
            Log.Information("DefaultFixer1 Processing");
        }
    }
}
