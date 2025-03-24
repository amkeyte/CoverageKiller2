using Microsoft.Office.Interop.Word;
using System;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a row in a table.
    /// Inherits from CKCellsRect to provide enumeration of the cells in the row.
    /// </summary>
    public class CKRow : CKCellsRect, IDOMObject
    {
        #region Fields

        ///// <summary>
        ///// Stores the last column index of the row.
        ///// </summary>
        //internal int _lastIndex;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRow"/> class using the specified table and a rectangular cell reference.
        /// The cell reference must represent exactly one row.
        /// </summary>
        /// <param name="table">The parent table that contains the row.</param>
        /// <param name="cellReference">A rectangular cell reference for one row.</param>
        /// <exception cref="ArgumentException">Thrown when the cell reference does not represent exactly one row.</exception>
        public CKRow(CKTable table, CKCellRefRect cellReference)
            : base(table, cellReference)
        {
            if (cellReference.Y1 != cellReference.Y2)
                throw new ArgumentException("A CKRow must represent exactly one row.", nameof(cellReference));

            //_lastIndex = cellReference.X2;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the cells in this row.
        /// Since CKRow is a specialized CKCellsRect, it can be used directly.
        /// </summary>
        public CKCells Cells => this;

        /// <summary>
        /// Gets the one-based row index.
        /// Since the cell reference is rectangular and represents a single row, this returns the Y1 coordinate.
        /// </summary>
        public int Index => ((CKCellRefRect)CellRef).Y1;

        /// <summary>
        /// Gets a CKRange that spans the entire row.
        /// This range is defined from the start of the first cell to the end of the last cell in the row.
        /// </summary>
        public CKRange Range
        {
            get
            {
                // Ensure there is at least one cell.
                if (!Cells.Any())
                    throw new InvalidOperationException("The row contains no cells.");

                CKCell firstCell = Cells.First();
                CKCell lastCell = Cells.Last();
                return Document.Range(firstCell.Start, lastCell.End, this);
            }
        }

        /// <summary>
        /// Gets the CKDocument associated with this row.
        /// This is derived from the parent table.
        /// </summary>
        public CKDocument Document => Table.Document;

        /// <summary>
        /// Gets the Word application instance managing the document.
        /// This is derived from the parent table.
        /// </summary>
        public Application Application => Table.Application;

        #endregion

        #region IDOMObject Members

        /// <summary>
        /// Gets a value indicating whether this row is dirty.
        /// A row is dirty if its parent table is dirty or if any cell in the row is dirty.
        /// </summary>
        public bool IsDirty => Table.IsDirty || Cells.Any(c => c.IsDirty);

        /// <summary>
        /// Gets a value indicating whether this row is orphaned.
        /// A row is orphaned if any of its underlying cell COM objects are no longer valid.
        /// </summary>
        public bool IsOrphan => Cells.Any(c => c.IsOrphan);

        #endregion
    }
}
