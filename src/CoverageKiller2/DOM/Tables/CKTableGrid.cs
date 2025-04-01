
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;

[assembly: InternalsVisibleTo("CoverageKiller2_Tests")]
namespace CoverageKiller2.DOM.Tables
{
    public class CKTableGrid
    {
        private static Dictionary<CKRange, CKTableGrid> _tableGrids = new Dictionary<CKRange, CKTableGrid>();

        private Word.Table _table;
        private GridCell[,] _grid;

        // 🍽 Shared across all internal grid ops
        internal int RowCount { get; set; }
        internal int ColCount { get; set; }

        public static CKTableGrid GetInstance(Word.Table table)
        {
            var tableRange = new CKRange(table.Range);

            _tableGrids.Keys.Where(r => r.IsOrphan).ToList()
                .ForEach(r => _tableGrids.Remove(r));

            if (_tableGrids.TryGetValue(tableRange, out CKTableGrid grid))
            {
                return grid;
            }

            grid = new CKTableGrid(table);
            _tableGrids.Add(tableRange, grid);

            return grid;
        }

        private CKTableGrid(Word.Table table)
        {
            _table = table;
            BuildGrid();
        }

        private void BuildGrid()
        {
            // 🧠 We store these once and reuse them — like a healthy trauma response.
            RowCount = _table.Rows.Count;
            ColCount = GetMaxColumns(_table);
            _grid = new GridCell[RowCount, ColCount];

            // 🧟 Looping all cells like it's a Word zombie apocalypse.
            foreach (Word.Cell cell in _table.Range.Cells)
            {
                int row = cell.RowIndex - 1;
                int col = cell.ColumnIndex - 1;

                if (_grid[row, col] != null)
                    continue;

                var (rowSpan, colSpan) = GetCellSpan(cell);

                for (int r = 0; r < rowSpan; r++)
                {
                    for (int c = 0; c < colSpan; c++)
                    {
                        int targetRow = row + r;
                        int targetCol = col + c;

                        if (targetRow < RowCount && targetCol < ColCount)
                        {
                            bool isMaster = (r == 0 && c == 0);
                            _grid[targetRow, targetCol] = new GridCell(cell, targetRow, targetCol, isMaster);
                        }
                    }
                }
            }
        }

        private int GetMaxColumns(Word.Table table)
        {
            int maxCol = 0;

            foreach (Word.Cell cell in table.Range.Cells)
            {
                int colIndex = cell.ColumnIndex;
                if (colIndex > maxCol)
                    maxCol = colIndex;
            }

            return maxCol;
        }
        // ☢️ This method now respects sanity caps, avoiding infinite traversal of a Word fever dream.
        internal (int rowSpan, int colSpan) GetCellSpan(Word.Cell cell)
        {
            int startRow = cell.RowIndex;
            int startCol = cell.ColumnIndex;
            int rowSpan = 1;
            int colSpan = 1;

            bool IsSameCell(Word.Cell a, Word.Cell b)
                => a != null && b != null && a.Range.Start == b.Range.Start;

            for (int r = startRow + 1; r <= RowCount; r++)
            {
                try
                {
                    var testCell = _table.Cell(r, startCol);
                    if (!IsSameCell(testCell, cell)) break;
                    rowSpan++;
                }
                catch { break; }
            }

            for (int c = startCol + 1; c <= ColCount; c++)
            {
                try
                {
                    var testCell = _table.Cell(startRow, c);
                    if (!IsSameCell(testCell, cell)) break;
                    colSpan++;
                }
                catch { break; }
            }

            // 🛡 Safety log for suspicious spans
            if (rowSpan > 50 || colSpan > 50)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Span {rowSpan}x{colSpan} at cell ({startRow},{startCol}) may be corrupted or unusually large.");
            }

            return (rowSpan, colSpan);
        }



        public IEnumerable<GridCell> GetMasterCells()
        {
            return Enumerable.Range(0, _grid.GetLength(0))
                .SelectMany(row => Enumerable.Range(0, _grid.GetLength(1))
                    .Select(col => _grid[row, col]))
                .Where(cell => cell != null && cell.IsMasterCell);
        }

        public IEnumerable<GridCell> GetMasterCells(CKGridCellRef area)
        {
            for (int row = area.Y1; row <= area.Y2; row++)
            {
                for (int col = area.X1; col <= area.X2; col++)
                {
                    if (row < 0 || row >= RowCount || col < 0 || col >= ColCount)
                        continue;

                    var cell = _grid[row, col];
                    if (cell != null && cell.IsMasterCell)
                        //if null cell (how?) thee could be unexpected behavior
                        yield return cell;
                }
            }
        }
    }

    //public class GridCell
    //{
    //    public Word.Cell COMCell { get; private set; }
    //    public bool IsMasterCell { get; private set; }
    //    public int GridRow { get; private set; }
    //    public int GridCol { get; private set; }
    //    public RangeSnapshot Snapshot { get; private set; }

    //    public GridCell(Word.Cell cell, int gridRow, int gridCol, bool isMasterCell, RangeSnapshot snapshot)
    //    {
    //        COMCell = cell;
    //        GridRow = gridRow;
    //        GridCol = gridCol;
    //        IsMasterCell = isMasterCell;
    //        Snapshot = snapshot;
    //    }
    //}
}
