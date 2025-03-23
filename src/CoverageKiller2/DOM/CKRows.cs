using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    public class CKRows : IEnumerable<CKRow>, IDOMObject
    {

        private List<CKRow> _rows = new List<CKRow>();
        public override string ToString()
        {
            return $"CKRows[Count: {Count}]";
        }

        public CKRows(IEnumerable<CKRow> rows, IDOMObject parent)
        {
            _rows = rows.ToList();
            Parent = parent;
        }

        public int Count => _rows.Count;

        public CKDocument Document => throw new NotImplementedException();

        public Application Application => throw new NotImplementedException();

        public IDOMObject Parent { get; private set; }


        IDOMObject IDOMObject.Parent => throw new NotImplementedException();

        public bool IsOrphan => _rows.Any(r => r.IsOrphan);

        public bool IsDirty => throw new NotImplementedException();

        // Access a CKRow by its index (1-based index in Word)
        public CKRow this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return _rows[index];
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKRow> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
