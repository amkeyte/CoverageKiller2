using CoverageKiller2.DOM;
using System.Collections.Generic;

public class CellRefCoord : ICellRef<CKCell>
{
    public IEnumerable<int> WordCells { get; }

    public int GridX1 { get; }
    public int GridY1 { get; }
    public int GridX2 { get; }
    public int GridY2 { get; }

    public CellRefCoord(int x, int y, int cellRef)
    {
        GridX1 = x;
        GridY1 = y;
        GridX2 = x;
        GridY2 = y;
        WordCells = new List<int>() { cellRef };
    }
}