using System.Collections.Generic;

namespace CoverageKiller2.DOM
{
    public interface ICellRef<out T> where T : IDOMObject
    {
        CKTable Table { get; }
        IEnumerable<int> CellIndexes { get; }
        IDOMObject Parent { get; }
    }
}
