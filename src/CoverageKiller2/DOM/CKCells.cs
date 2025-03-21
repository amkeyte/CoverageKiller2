using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    public class CKCells : IEnumerable<CKCell>
    {

        private List<CKCell> _cells = new List<CKCell>();
        private List<CKCell> Items
        {
            get
            {
                if (IsDirty || Table.IsDirty) _cells = Table.CellItems(CellReference).ToList();
                return _cells;
            }
        }
        public CKTable Table { get; private set; }

        public CKCells(CKTable table, CKCellReference cellReference)
        {
            Table = table;
            CellReference = cellReference;
        }
        public CKCells(CKRange range)
        {
            var table = range.Tables.FirstOrDefault() ??
                throw new InvalidOperationException("Range does not contain a table");

            Table = table;
            CellReference = new CKCellReference(range);

        }
        public int Count => Items.Count;

        public bool IsDirty => Items.Any(c => c.IsDirty);

        public CKCellReference CellReference { get; private set; }

        public IEnumerator<CKCell> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }
        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of sections.");
                return Items.ElementAt(index);
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
