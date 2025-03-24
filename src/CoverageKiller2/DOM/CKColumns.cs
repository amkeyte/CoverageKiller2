using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    public class CKColumns : IEnumerable<CKColumn>, IDOMObject
    {
        List<CKColumn> _columns;

        public CKColumns(IEnumerable<CKColumn> columns, IDOMObject parent)
        {
            _columns = columns?.ToList() ?? throw new ArgumentNullException(nameof(columns));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }


        // Property to get the total number of columns
        public int Count => _columns.Count();

        /// <summary>
        /// Gets the CKDocument associated with these rows.
        /// This is derived from the Parent property.
        /// </summary>
        public CKDocument Document => Parent.Document;

        /// <summary>
        /// Gets the Word application managing the document.
        /// This is derived from the Parent property.
        /// </summary>
        public Application Application => Parent.Application;

        /// <summary>
        /// Gets the parent DOM object.
        /// </summary>
        public IDOMObject Parent { get; private set; }

        // Access a CKColumn by its current index (1-based index in Word)
        public CKColumn this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return _columns[index];
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKColumn> GetEnumerator()
        {
            for (int index = 1; index <= Count; index++)
            {
                yield return this[index];
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        /// <summary>
        /// Gets a value indicating whether any row in this collection is orphaned.
        /// A row is orphaned if its underlying COM objects are no longer valid.
        /// </summary>
        public bool IsOrphan => _columns.Any(r => r.IsOrphan);

        /// <summary>
        /// Gets a value indicating whether any row in this collection is dirty.
        /// </summary>
        public bool IsDirty => _columns.Any(r => r.IsDirty);
    }
}
