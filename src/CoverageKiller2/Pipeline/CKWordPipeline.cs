using CoverageKiller2.Pipeline.Config;
using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.Pipeline
{
    /// <summary>
    /// Represents a pipeline for processing Word documents using a collection of pipeline processes.
    /// </summary>
    public class CKWordPipeline : ICollection<CKWordPipelineProcess>
    {
        private readonly ICollection<CKWordPipelineProcess> _items = new List<CKWordPipelineProcess>();

        /// <summary>
        /// Gets the document associated with the pipeline.
        /// </summary>
        public CKDocument Document { get; }
        public CKDocument Template { get; private set; }
        public ProcessorConfig ProcessorConfig { get; private set; }

        /// <summary>
        /// Gets the number of items in the pipeline.
        /// </summary>
        public int Count => _items.Count;

        /// <summary>
        /// Gets a value indicating whether the pipeline is read-only.
        /// </summary>
        public bool IsReadOnly => _items.IsReadOnly;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKWordPipeline"/> class with the specified document.
        /// </summary>
        /// <param name="ckDoc">The CKDocument to be processed.</param>
        public CKWordPipeline(CKDocument ckDoc)
        {
            Document = ckDoc;
        }

        public CKWordPipeline(Dictionary<string, object> initVars)
        {
            Document = (CKDocument)initVars["ckDoc"];
            Template = (CKDocument)initVars["template"];
            ProcessorConfig = (ProcessorConfig)initVars["ProcessorConfig"];
        }

        /// <summary>
        /// Adds a process to the pipeline and assigns the document to the process.
        /// </summary>
        /// <param name="item">The pipeline process to add.</param>
        public void Add(CKWordPipelineProcess item)
        {
            item.CKDoc = Document;
            item.ProcessorConfig = ProcessorConfig;
            if (item.Template is null) item.Template = Template;
            _items.Add(item);
        }

        /// <summary>
        /// Clears all processes from the pipeline.
        /// </summary>
        public void Clear()
        {
            _items.Clear();
        }

        /// <summary>
        /// Determines whether the pipeline contains a specific process.
        /// </summary>
        /// <param name="item">The process to locate in the pipeline.</param>
        /// <returns><c>true</c> if the process is found; otherwise, <c>false</c>.</returns>
        public bool Contains(CKWordPipelineProcess item)
        {
            return _items.Contains(item);
        }

        /// <summary>
        /// Copies the elements of the pipeline to an array, starting at the specified index.
        /// </summary>
        /// <param name="array">The array to copy the pipeline elements into.</param>
        /// <param name="arrayIndex">The zero-based index in the array where copying begins.</param>
        public void CopyTo(CKWordPipelineProcess[] array, int arrayIndex)
        {
            _items.CopyTo(array, arrayIndex);
        }

        /// <summary>
        /// Removes a specific process from the pipeline.
        /// </summary>
        /// <param name="item">The process to remove.</param>
        /// <returns><c>true</c> if the process was successfully removed; otherwise, <c>false</c>.</returns>
        public bool Remove(CKWordPipelineProcess item)
        {
            return _items.Remove(item);
        }

        /// <summary>
        /// Returns an enumerator that iterates through the pipeline processes.
        /// </summary>
        /// <returns>An enumerator for the pipeline processes.</returns>
        public IEnumerator<CKWordPipelineProcess> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_items).GetEnumerator();
        }

        /// <summary>
        /// Runs all the processes in the pipeline sequentially.
        /// </summary>
        internal void Run()
        {
            foreach (var item in _items)
            {
                item.Process();
            }
        }
    }
}
