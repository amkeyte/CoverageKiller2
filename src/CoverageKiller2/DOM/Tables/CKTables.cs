using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Represents a collection of <see cref="CKTable"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0001
    /// </remarks>
    public class CKTables : ACKRangeCollection, IEnumerable<CKTable>
    {
        private Word.Tables _comTables;
        private Base1List<CKTable> _cachedTables_1 = new Base1List<CKTable>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTables"/> class from the specified parent range.
        /// </summary>
        /// <param name="collection">The Word tables collection.</param>
        /// <param name="parent">The parent is the Document that owns the tables.  </param>
        /// <remarks>In CSTO, the parent of a Tables collection is ALWAYS the document,
        /// regardles of the range it was pulled from</remarks>
        public CKTables(Word.Tables collection, CKDocument parent) : base(parent)
        {

            if (!parent.Matches(collection.Parent))
                throw new ArgumentException("collection and parent must have the same document.");

            _comTables = collection;
        }

        /// <inheritdoc/>
        public override int Count => this.PingPong(() => TablesList_1.Count);

        /// <inheritdoc/>
        public override void Clear() => _cachedTables_1.Clear();

        /// <inheritdoc/>
        public override bool IsOrphan => throw new NotImplementedException();

        /// <inheritdoc/>
        protected override bool CheckDirtyFor()
        {
            this.PingPong();
            // Placeholder logic: needs refinement
            return false;
        }

        private Word.Tables COMTables => this.PingPong(() => _comTables);

        private Base1List<CKTable> TablesList_1
        {
            get
            {
                this.Ping(msg: Document.FileName);

                if (!_cachedTables_1.Any() || IsDirty)
                {
                    _cachedTables_1.Clear();
                    for (int i = 1; i <= COMTables.Count; i++)
                    {
                        _cachedTables_1.Add(new CKTable(COMTables[i], this));
                    }
                    IsDirty = false;
                }

                this.Pong();
                return _cachedTables_1;
            }
        }

        /// <summary>
        /// Deletes the table at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the table to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if index is out of bounds.</exception>
        public void Delete(int index)
        {

            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");

            // Grab the CKTable at that index
            var ckTable = this[index];
            Delete(ckTable);

        }

        /// <summary>
        /// Deletes the specified CKTable from the document.
        /// </summary>
        /// <param name="table">The table to delete.</param>
        /// <exception cref="ArgumentNullException">Thrown if table is null.</exception>
        public void Delete(CKTable table)
        {

            if (table == null)
                throw new ArgumentNullException(nameof(table));

            // Call delete on the underlying Word.Table
            table.COMTable.Delete();

            IsDirty = true; // Force reload on next access
            _cachedTables_1.Clear(); // Clear cache immediately maybe maybe not. 

        }

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">One-based index of the table.</param>
        /// <returns>The corresponding <see cref="CKTable"/>.</returns>
        public CKTable this[int index]
        {
            get
            {
                this.Ping(Document.FileName);
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                var result = TablesList_1[index];
                this.Pong();
                return result;
            }
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            this.Ping(Document.FileName);

            if (obj is CKTable table)
            {
                int index = TablesList_1.IndexOf(table);
                this.Pong();
                return index;
            }

            this.Pong();
            return -1;
        }

        /// <summary>
        /// Adds a new <see cref="CKTable"/> at the specified range with the given number of rows and columns.
        /// </summary>
        /// <param name="insertAt">The <see cref="CKRange"/> at which to insert the table.</param>
        /// <param name="numRows">The number of rows in the new table.</param>
        /// <param name="numColumns">The number of columns in the new table.</param>
        /// <returns>The newly created <see cref="CKTable"/> instance.</returns>
        /// <exception cref="ArgumentNullException"/>
        /// <exception cref="ArgumentOutOfRangeException"/>
        /// <remarks>
        /// Version: CK2.00.01.0006
        /// </remarks>
        public CKTable Add(CKRange insertAt, int numRows, int numColumns)
        {
            this.Ping(Document.FileName);

            if (insertAt == null) throw new ArgumentNullException(nameof(insertAt));
            if (numRows < 1) throw new ArgumentOutOfRangeException(nameof(numRows));
            if (numColumns < 1) throw new ArgumentOutOfRangeException(nameof(numColumns));
            if (insertAt.Start != insertAt.End)
                throw new ArgumentException($"{nameof(insertAt)} must be collapsed.");

            var wordTable = COMTables.Add(insertAt.COMRange, numRows, numColumns);

            IsDirty = true;
            return this.Pong(() => new CKTable(wordTable, this));
        }

        /// <summary>
        /// Returns the <see cref="CKTable"/> that owns the given <see cref="Word.Cell"/>, if present.
        /// </summary>
        /// <param name="cell">A Word cell to search for.</param>
        /// <returns>The owning <see cref="CKTable"/>.</returns>
        /// <exception cref="ArgumentOutOfRangeException"/>
        /// <remarks>
        /// Version: CK2.00.01.0004
        /// </remarks>
        internal CKTable ItemOf(Word.Cell cell)
        {
            this.Ping(Document.FileName);

            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var ckTable = TablesList_1.FirstOrDefault(t => t.Contains(cell))
                ?? throw new ArgumentOutOfRangeException(nameof(cell), "Cell is not contained in any known table.");

            this.Pong();
            return ckTable;
        }

        /// <inheritdoc/>
        public override string ToString() => $"CKTables [Count: {Count}]";

        /// <inheritdoc/>
        public IEnumerator<CKTable> GetEnumerator() => TablesList_1.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
