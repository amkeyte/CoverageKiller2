namespace CoverageKiller2
{
    public abstract class CKWordPipelineProcess
    {
        public abstract void Process();
        public CKDocument CKDoc { get; set; }
    }
}
