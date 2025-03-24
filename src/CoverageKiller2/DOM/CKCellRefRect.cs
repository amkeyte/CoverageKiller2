using System;

namespace CoverageKiller2.DOM
{


    /// <summary>
    /// Represents an immutable cell reference defined by its starting and ending coordinates.
    /// Coordinates are one-based, with X representing the column and Y the row.
    /// </summary>
    public class CKCellRefRect : ICellRef
    {
        /// <summary>
        /// Gets the starting column index.
        /// </summary>
        public int X1 { get; }

        /// <summary>
        /// Gets the starting row index.
        /// </summary>
        public int Y1 { get; }

        /// <summary>
        /// Gets the ending column index.
        /// </summary>
        public int X2 { get; }

        /// <summary>
        /// Gets the ending row index.
        /// </summary>
        public int Y2 { get; }

        public string Mode => nameof(CKCellRefRect);

        public int Start => throw new NotSupportedException();

        public int End => throw new NotImplementedException();



        /// <summary>
        /// Private constructor used by the factory methods.
        /// </summary>
        private CKCellRefRect(int x1, int y1, int x2, int y2)
        {
            if (x1 < 1 || y1 < 1)
                throw new ArgumentOutOfRangeException("Coordinates must be at least 1.");
            if (x1 > x2 || y1 > y2)
                throw new ArgumentException("Coordinates must be provided in increasing order: x1 <= x2 and y1 <= y2.");

            X1 = x1;
            Y1 = y1;
            X2 = x2;
            Y2 = y2;
        }

        /// <summary>
        /// Creates a cell reference for a single cell.
        /// </summary>
        /// <param name="x">The column index (one-based) of the cell.</param>
        /// <param name="y">The row index (one-based) of the cell.</param>
        /// <returns>A CKCellRef that refers to the single cell at (x, y).</returns>
        public static CKCellRefRect ForCell(int x, int y)
        {
            return new CKCellRefRect(x, y, x, y);
        }

        /// <summary>
        /// Creates a cell reference for an entire row.
        /// </summary>
        /// <param name="rowIndex">The one-based row index.</param>
        /// <param name="columnCount">The total number of columns in the table.</param>
        /// <returns>A CKCellRef that covers the entire row.</returns>
        public static CKCellRefRect ForRow(int rowIndex, int columnCount)
        {
            if (rowIndex < 1)
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be at least 1.");
            if (columnCount < 1)
                throw new ArgumentOutOfRangeException(nameof(columnCount), "Column count must be at least 1.");

            return new CKCellRefRect(1, rowIndex, columnCount, rowIndex);
        }

        /// <summary>
        /// Creates a cell reference for an entire column.
        /// </summary>
        /// <param name="columnIndex">The one-based column index.</param>
        /// <param name="rowCount">The total number of rows in the table.</param>
        /// <returns>A CKCellRef that covers the entire column.</returns>
        public static CKCellRefRect ForColumn(int columnIndex, int rowCount)
        {
            if (columnIndex < 1)
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be at least 1.");
            if (rowCount < 1)
                throw new ArgumentOutOfRangeException(nameof(rowCount), "Row count must be at least 1.");

            return new CKCellRefRect(columnIndex, 1, columnIndex, rowCount);
        }

        /// <summary>
        /// Creates a cell reference for a range of cells.
        /// </summary>
        /// <param name="x1">The starting column index.</param>
        /// <param name="y1">The starting row index.</param>
        /// <param name="x2">The ending column index.</param>
        /// <param name="y2">The ending row index.</param>
        /// <returns>A CKCellRef covering the specified range.</returns>
        public static CKCellRefRect ForRectangle(int x1, int y1, int x2, int y2)
        {
            return new CKCellRefRect(x1, y1, x2, y2);
        }

        public override string ToString()
        {
            return $"CKCellRef: ({X1}, {Y1}) to ({X2}, {Y2})";
        }

        public CKCellRefRect ToRect()
        {
            throw new NotImplementedException();
        }
    }
}
