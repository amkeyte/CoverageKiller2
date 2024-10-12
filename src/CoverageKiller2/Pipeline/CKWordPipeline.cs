using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.Pipeline
{
    public class CKWordPipeline : ICollection<CKWordPipelineProcess>
    {
        public int Count => _items.Count;

        public bool IsReadOnly => _items.IsReadOnly;

        public CKDocument Document { get; }

        private readonly ICollection<CKWordPipelineProcess> _items = new List<CKWordPipelineProcess>();

        public CKWordPipeline(CKDocument ckDoc)
        {
            Document = ckDoc;
        }

        public void Add(CKWordPipelineProcess item)
        {
            _items.Add(item);
            item.CKDoc = Document;
        }

        public void Clear()
        {
            _items.Clear();
        }

        public bool Contains(CKWordPipelineProcess item)
        {
            return _items.Contains(item);
        }

        public void CopyTo(CKWordPipelineProcess[] array, int arrayIndex)
        {
            _items.CopyTo(array, arrayIndex);
        }

        public IEnumerator<CKWordPipelineProcess> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        public bool Remove(CKWordPipelineProcess item)
        {
            return _items.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_items).GetEnumerator();
        }

        internal void Run()
        {
            foreach (CKWordPipelineProcess item in _items)
            {
                item.Process();
            }
        }
    }
}
