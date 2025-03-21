using CoverageKiller2.DOM;
using CoverageKiller2.Pipeline.Config;

namespace CoverageKiller2.Pipeline
{
    /// <summary>
    /// Represents an abstract base class for processing Word documents in the pipeline.
    /// </summary>
    public abstract class CKWordPipelineProcess
    {
        /// <summary>
        /// Processes the document.
        /// </summary>
        public abstract void Process();

        /// <summary>
        /// Gets or sets the CKDocument associated with this process.
        /// </summary>
        public CKDocument CKDoc { get; set; }
        public ProcessorConfig ProcessorConfig { get; internal set; }
        public CKDocument Template { get; internal set; }
    }
}
