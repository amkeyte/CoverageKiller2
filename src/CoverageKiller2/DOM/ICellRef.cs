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
}
