using CoverageKiller2.Logging;

namespace CoverageKiller2.Pipeline.Processes
{
    internal class SEA2025Fixer : CKWordPipelineProcess
    {
        public Tracer Tracer { get; } = new Tracer(typeof(SEA2025Fixer));
        public override void Process()
        {
            throw new System.NotImplementedException();
        }
        public SEA2025Fixer()
        {

        }

    }
}
