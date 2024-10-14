namespace CoverageKiller2.Pipeline
{
    public abstract class CKWordPipelineProcess
    {
        public abstract void Process();
        public CKDocument CKDoc { get; set; }
    }
}
