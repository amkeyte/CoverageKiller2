using System;

namespace CoverageKiller2.DOM
{
    public interface ICellRef
    {
        string Mode { get; }

        int Start { get; }

        int End { get; }

        /// <summary>
        /// Gets the starting column index.
        /// </summary>
        int X1 { get; }

        /// <summary>
        /// Gets the starting row index.
        /// </summary>
        int Y1 { get; }

        /// <summary>
        /// Gets the ending column index.
        /// </summary>
        int X2 { get; }

        /// <summary>
        /// Gets the ending row index.
        /// </summary>
        int Y2 { get; }
    }

    public class CKCellRefLinear : ICellRef
    {
        private CKCellRefLinear(int start, int end)
        {
            Start = start;
            End = end;
        }

        /// <summary>
        /// Gets the mode of this cell reference.
        /// </summary>
        public string Mode => nameof(CKCellRefLinear);

        /// <summary>
        /// In linear mode, Start represents the first cell's one-based index in the table's Cells collection.
        /// </summary>
        public int Start { get; }

        /// <summary>
        /// In linear mode, End represents the last cell's one-based index in the table's Cells collection.
        /// </summary>
        public int End { get; }

        // The rectangular properties are not supported in linear mode.
        public int X1 => throw new NotSupportedException();
        public int Y1 => throw new NotSupportedException();
        public int X2 => throw new NotSupportedException();
        public int Y2 => throw new NotSupportedException();

        /// <summary>
        /// Creates a CKCellLinear instance representing a range of cells
        /// from the table's Cells collection, where the indices are one-based.
        /// </summary>
        /// <param name="start">The starting cell index (one-based).</param>
        /// <param name="end">The ending cell index (one-based, must be >= start).</param>
        /// <returns>A CKCellLinear instance for the specified range.</returns>
        public static CKCellRefLinear ForCells(int start, int end)
        {
            if (start < 1)
                throw new ArgumentOutOfRangeException(nameof(start), "Start must be at least 1.");
            if (end < start)
                throw new ArgumentException("End must be greater than or equal to start.", nameof(end));

            return new CKCellRefLinear(start, end);
        }
        public static CKCellRefLinear ForCells(CKRange range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            // Assuming that range.COMRange.Cells gives the cells for the range.
            int cellCount = range.COMRange.Cells.Count;
            return new CKCellRefLinear(1, cellCount);
        }

        /// <summary>
        /// Creates a CKCellLinear instance representing a single cell in the table's Cells collection.
        /// </summary>
        /// <param name="index">The one-based index of the cell.</param>
        /// <returns>A CKCellLinear instance for the single cell.</returns>
        public static CKCellRefLinear ForCell(int index)
        {
            return ForCells(index, index);
        }

        public override string ToString()
        {
            return $"CKCellLinear: Cells {Start} to {End}";
        }
    }


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
    }
}
