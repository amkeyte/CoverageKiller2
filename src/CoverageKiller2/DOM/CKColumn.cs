namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a single column in a Word table, implemented as a rectangular cell collection.
    /// This column is defined by a cell reference that must represent exactly one column.
    /// </summary>
    public class CKColumn : CKCellsRect, IDOMObject
    {
        public CKColumn(CKTable table, ICellRef<CKCellsRect> cellReference) : base(table, cellReference)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKColumn"/> class using the specified table and cell reference.
        /// The cell reference must represent exactly one column.
        /// </summary>
        /// <param name="table">The parent CKTable that contains this column.</param>
        /// <param name="cellRef">A rectangular cell reference representing one column (X1 must equal X2).</param>
        /// <exception cref="ArgumentNullException">Thrown when table or cellRef is null.</exception>
        /// <exception cref="ArgumentException">Thrown when the cell reference does not represent exactly one column.</exception>
        //public CKColumn(CKTable table, CKGridCellRefRect cellRef)
        //    : base(table, cellRef)
        //{
        //    if (table == null)
        //        throw new ArgumentNullException(nameof(table));
        //    if (cellRef == null)
        //        throw new ArgumentNullException(nameof(cellRef));
        //    if (cellRef.X1 != cellRef.X2)
        //        throw new ArgumentException("A CKColumn must represent exactly one column.", nameof(cellRef));

        //    Table = table;
        //}

        /// <summary>
        /// Gets the one-based column index.
        /// </summary>
        //public int Index => ((CKGridCellRefRect)CellRef).X1;

        ///// <summary>
        ///// Gets the parent DOM object. For CKColumn, the parent is the CKTable.
        ///// </summary>
        //public IDOMObject Parent { get; private set; }





        /// <summary>
        /// Returns a string representation of this CKColumn.
        /// </summary>
        //public override string ToString() => $"CKColumn[Index: {Index}, Cells: {Count}]";
    }
}
