using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents an arbitrary collection of CKCell objects.
    /// </summary>
    public abstract class CKCells : IEnumerable<CKCell>, IDOMObject
    {
        protected List<CKCell> _cells;
        public CKTable Table { get; protected set; }
        public ICellRef CellRef { get; protected set; }

        protected CKCells(CKTable table, ICellRef cellReference)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            CellRef = cellReference ?? throw new ArgumentNullException(nameof(cellReference));
            _cells = BuildCells().ToList();
        }

        /// <summary>
        /// Derived classes implement this method to build the cell collection.
        /// </summary>
        /// <returns>An enumerable of CKCell objects.</returns>
        protected abstract IEnumerable<CKCell> BuildCells();

        public int Count => _cells.Count;

        public CKDocument Document => Table.Document;
        public Word.Application Application => Table.Application;
        public IDOMObject Parent => Table;
        public bool IsDirty => Table.IsDirty || _cells.Any(c => c.IsDirty);
        public bool IsOrphan => Table.IsOrphan;

        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _cells[index - 1];
            }
        }

        public IEnumerator<CKCell> GetEnumerator() => _cells.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// Represents a collection of CKCell objects built from a linear cell reference.
    /// </summary>
    public class CKCellsLinear : CKCells
    {
        public CKCellsLinear(CKTable table, ICellRef cellReference)
            : base(table, cellReference)
        {
            if (cellReference.Mode != nameof(CKCellRefLinear))
                throw new NotSupportedException($"For CKCellsLinear, cellReference must be of type {nameof(CKCellRefLinear)}.");
        }
        /// <summary>
        /// Constructs a CKCells collection from a CKRange.
        /// Assumes the first table in the range.
        /// </summary>
        public CKCellsLinear(CKRange range) : this(range.Tables[1], CKCellRefLinear.ForCells(range))
        {
        }

        protected override IEnumerable<CKCell> BuildCells()
        {
            // In linear mode, iterate over the table's Cells based on sequential indices.
            for (int i = CellRef.Start; i <= CellRef.End; i++)
            {
                // Table.COMTable.Cells is assumed to be a one-based collection.
                yield return new CKCell(Table.COMTable.Range.Cells[i]);
            }
        }
    }

    /// <summary>
    /// Represents a collection of CKCell objects that form a contiguous rectangular grid.
    /// </summary>
    public class CKCellsRect : CKCells
    {
        public CKCellsRect(CKTable table, ICellRef cellReference)
            : base(table, cellReference)
        {
            if (cellReference.Mode != nameof(CKCellRefRect))
                throw new NotSupportedException($"For CKCellsRect, cellReference must be of type {nameof(CKCellRefRect)}.");
        }

        protected override IEnumerable<CKCell> BuildCells()
        {
            // In rectangular mode, iterate over rows and columns.
            for (int row = CellRef.Y1; row <= CellRef.Y2; row++)
            {
                for (int col = CellRef.X1; col <= CellRef.X2; col++)
                {
                    yield return new CKCell(Table.COMTable.Cell(row, col));
                }
            }
        }
    }
}
