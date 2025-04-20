using CoverageKiller2.Logging;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Represents a rectangular cell reference within a CKTableGrid.
    /// Coordinates are 1 based and inclusive.
    /// </summary>
    public readonly struct CKGridCellRef
    {
        /// <summary>
        /// The starting column 
        /// </summary>
        public int ColMin { get; }

        /// <summary>
        /// The starting row 
        /// </summary>
        public int RowMin { get; }

        /// <summary>
        /// The ending column 
        /// </summary>
        public int ColMax { get; }

        /// <summary>
        /// The ending row 
        /// </summary>
        public int RowMax { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKGridCellRef"/> struct.
        /// </summary>
        /// <param name="cMin">Start column (inclusive, zero-based).</param>
        /// <param name="rMin">Start row (inclusive, zero-based).</param>
        /// <param name="cMax">End column (inclusive, zero-based).</param>
        /// <param name="rMax">End row (inclusive, zero-based).</param>
        public CKGridCellRef(int rMin, int cMin, int rMax, int cMax)
        {
            ColMin = cMin;
            RowMin = rMin;
            ColMax = cMax;
            RowMax = rMax;
            this.PingPong();


        }

        /// <summary>
        /// Returns a string representation of the grid cell reference.
        /// </summary>
        public override string ToString() => $"GridCellRef [({ColMin}, {RowMin}) - ({ColMax}, {RowMax})]";
    }
}
