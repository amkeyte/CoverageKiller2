using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents an arbitrary collection of CKCell objects.
    /// This collection may be non-rectangular.
    /// </summary>
    public class CKCells : IEnumerable<CKCell>
    {
        protected List<CKCell> _cells;

        /// <summary>
        /// Gets the parent table for these cells.
        /// </summary>
        public CKTable Table { get; protected set; }

        /// <summary>
        /// Gets the cell reference defining the numeric boundaries.
        /// </summary>
        public ICellRef CellRef { get; protected set; }

        public CKCells(CKRange range) : this(range.Tables[1], CKCellRefLinear.ForCells(range))
        {

        }

        public CKCells(CKTable table, ICellRef cellReference)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            if (cellReference == null)
                throw new ArgumentNullException(nameof(cellReference));

            CellRef = cellReference;
            Table = table;
            _cells = BuildCells().ToList();
        }

        /// <summary>
        /// Builds the collection of cells based on the CellReference.
        /// In a non-rectangular collection, not every cell in the encompassing rectangle is necessarily valid.
        /// </summary>
        protected virtual IEnumerable<CKCell> BuildCells()
        {
            _cells = new List<CKCell>();

            if (CellRef.Mode == nameof(CKCellRefLinear))
            {

                // Table.COMTable.Cells is assumed to be a one-based collection.
                for (int i = CellRef.Start; i <= CellRef.End; i++)
                {
                    yield return new CKCell(Table.COMTable.Range.Cells[i]);

                }
            }
        }

        public int Count => _cells.Count;


        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _cells[index - 1];
            }
        }

        public IEnumerator<CKCell> GetEnumerator()
        {
            return _cells.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    /// <summary>
    /// Represents a collection of CKCell objects that form a contiguous rectangular grid.
    /// </summary>
    public class CKCellsRectangle : CKCells
    {
        public CKCellsRectangle(CKTable table, ICellRef cellReference)
            : base(table, cellReference)
        {
            if (cellReference.Mode != nameof(CKCellRefRect))
                throw new NotSupportedException($"{nameof(cellReference)} must be  {nameof(CKCellRefRect)}");

            _cells = BuildCells().ToList();
        }

        protected override IEnumerable<CKCell> BuildCells()
        {
            // For rectangular cell references, iterate over rows and columns.
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
