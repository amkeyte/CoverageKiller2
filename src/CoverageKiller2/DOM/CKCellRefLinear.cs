using System;

namespace CoverageKiller2.DOM
{
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

        public CKCellRefRect ToRect()
        {
            throw new NotImplementedException();
        }
    }
}
