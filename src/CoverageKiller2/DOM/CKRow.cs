using Microsoft.Office.Interop.Word;
using System;
using System.Linq;


namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a row in a table. Inherits from a rectangular cell collection,
    /// so that all the cells in that row can be enumerated.
    /// </summary>
    public class CKRow : CKCellsRect, IDOMObject
    {
        // Stores the last column index of the row (if needed for additional logic)
        internal int _lastIndex;

        /// <summary>
        /// Constructs a CKRow from a table and a rectangular cell reference.
        /// The cell reference must represent exactly one row.
        /// </summary>
        /// <param name="table">The parent table.</param>
        /// <param name="cellReference">A rectangular cell reference for one row.</param>
        public CKRow(CKTable table, CKCellRefRect cellReference) : base(table, cellReference)
        {
            if (cellReference.Y1 != cellReference.Y2)
                throw new ArgumentException("A CKRow must represent exactly one row.", nameof(cellReference));
            // Store the last column index from the cell reference.
            _lastIndex = cellReference.X2;
        }

        /// <summary>
        /// Gets the cells of this row.
        /// </summary>
        public CKCells Cells => this;

        /// <summary>
        /// Gets the one-based row index.
        /// Since the cell reference is rectangular and represents a single row,
        /// this returns the Y1 coordinate.
        /// </summary>
        public int Index => ((CKCellRefRect)CellRef).Y1;

        public CKRange Range => Document.Range(Cells.First().Start, Cells.Last().End, this);

        public CKDocument Document => Parent.Document;

        public Application Application => Parent.Application;

        // need to get appropriate CKRows object
        public IDOMObject Parent => throw new NotImplementedException();

        public bool IsDirty => throw new NotImplementedException();
        /// <summary>
        /// Gets a value indicating whether this CKDocument no longer has a valid COMDocument reference.
        /// This becomes true if the document is closed or the COM object has been released.
        /// </summary>
        public bool IsOrphan => Cells.Any(c => c.IsOrphan);
    }
}
