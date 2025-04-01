namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Represents a rectangular cell reference within a CKTableGrid.
    /// Coordinates are zero-based and inclusive.
    /// </summary>
    public readonly struct CKGridCellRef
    {
        /// <summary>
        /// The starting column (zero-based).
        /// </summary>
        public int X1 { get; }

        /// <summary>
        /// The starting row (zero-based).
        /// </summary>
        public int Y1 { get; }

        /// <summary>
        /// The ending column (zero-based).
        /// </summary>
        public int X2 { get; }

        /// <summary>
        /// The ending row (zero-based).
        /// </summary>
        public int Y2 { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKGridCellRef"/> struct.
        /// </summary>
        /// <param name="x1">Start column (inclusive, zero-based).</param>
        /// <param name="y1">Start row (inclusive, zero-based).</param>
        /// <param name="x2">End column (inclusive, zero-based).</param>
        /// <param name="y2">End row (inclusive, zero-based).</param>
        public CKGridCellRef(int x1, int y1, int x2, int y2)
        {
            X1 = x1;
            Y1 = y1;
            X2 = x2;
            Y2 = y2;
        }

        /// <summary>
        /// Returns a string representation of the grid cell reference.
        /// </summary>
        public override string ToString() => $"GridCellRef [({X1}, {Y1}) - ({X2}, {Y2})]";
    }
}
