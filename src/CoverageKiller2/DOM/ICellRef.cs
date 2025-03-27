using System.Collections.Generic;

namespace CoverageKiller2.DOM
{
    public interface ICellRef<out T> where T : IDOMObject
    {
        // Removed table-dependent logic from interface
        // Conversion will be handled within CKTable
        IEnumerable<int> WordCells { get; }
        int GridX1 { get; }
        int GridY1 { get; }
        int GridX2 { get; }
        int GridY2 { get; }
    }
}
