using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of CKRow objects in a Word table.
    /// This collection is part of the DOM hierarchy and implements IDOMObject.
    /// </summary>
    public class CKRows : IEnumerable<CKRow>, IDOMObject
    {
        #region Fields

        private List<CKRow> _rows = new List<CKRow>();

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRows"/> class with the specified rows and parent.
        /// </summary>
        /// <param name="rows">The enumerable collection of CKRow objects.</param>
        /// <param name="parent">The parent DOM object associated with these rows (typically a CKTable or CKDocument).</param>
        /// <exception cref="ArgumentNullException">Thrown when rows or parent is null.</exception>
        public CKRows(IEnumerable<CKRow> rows, IDOMObject parent)
        {
            if (rows == null)
                throw new ArgumentNullException(nameof(rows));
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));

            _rows = rows.ToList();
            Parent = parent;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the number of rows in this collection.
        /// </summary>
        public int Count => _rows.Count;

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

        /// <summary>
        /// Gets a value indicating whether any row in this collection is orphaned.
        /// A row is orphaned if its underlying COM objects are no longer valid.
        /// </summary>
        public bool IsOrphan => _rows.Any(r => r.IsOrphan);

        /// <summary>
        /// Gets a value indicating whether any row in this collection is dirty.
        /// </summary>
        public bool IsDirty => _rows.Any(r => r.IsDirty);

        /// <summary>
        /// Gets the CKRow at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the row to retrieve.</param>
        /// <returns>The CKRow at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if index is less than 1 or greater than Count.</exception>
        public CKRow this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");
                return _rows[index - 1];
            }
        }

        #endregion

        #region IEnumerable Implementation

        /// <summary>
        /// Returns an enumerator that iterates through the CKRow collection.
        /// </summary>
        /// <returns>An enumerator for the collection of CKRow objects.</returns>
        public IEnumerator<CKRow> GetEnumerator()
        {
            // Using one-based indexing.
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the CKRow collection.
        /// </summary>
        /// <returns>An enumerator for the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        #endregion
    }
}
