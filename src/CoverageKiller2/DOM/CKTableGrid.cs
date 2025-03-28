using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKTableGrid
    {
        private static Dictionary<CKRange, CKTableGrid> _tableGrids = new Dictionary<CKRange, CKTableGrid>();

        private Word.Table _table;
        private GridCell[,] _grid;

        public int RowCount { get; private set; }
        public int ColCount { get; private set; }

        public static CKTableGrid GetInstance(Word.Table table)
        {
            var tableRange = new CKRange(table.Range);

            //purge orphan references as temp fix for orphan RCP bug.
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
            int rowCount = _table.Rows.Count; //change to GetMaxColumns if needed.
            int colCount = GetMaxColumns(_table);
            RowCount = rowCount;
            ColCount = colCount;
            _grid = new GridCell[rowCount, colCount];

            for (int i = 1; i <= rowCount; i++)
            {
                Word.Row row = _table.Rows[i];
                int currentGridCol = 0;

                for (int j = 1; j <= row.Cells.Count; j++)
                {
                    Word.Cell cell = row.Cells[j];
                    while (currentGridCol < colCount && _grid[i - 1, currentGridCol] != null)
                        currentGridCol++;

                    if (currentGridCol >= colCount) break;

                    (int rowSpan, int colSpan) = GetCellSpan(cell);

                    for (int r = 0; r < rowSpan; r++)
                    {
                        for (int c = 0; c < colSpan; c++)
                        {
                            int targetRow = (i - 1) + r;
                            int targetCol = currentGridCol + c;
                            if (targetRow < rowCount && targetCol < colCount)
                            {
                                bool isMaster = (r == 0 && c == 0);
                                _grid[targetRow, targetCol] = new GridCell(cell, targetRow, targetCol, isMaster);
                            }
                        }
                    }

                    currentGridCol += colSpan;
                }
            }
        }

        private int GetMaxColumns(Word.Table table)
        {
            int max = 0;
            foreach (Word.Row row in table.Rows)
            {
                if (row.Cells.Count > max)
                    max = row.Cells.Count;
            }
            return max;
        }

        private (int rowSpan, int colSpan) GetCellSpan(Word.Cell cell)
        {
            // TODO: Proper merged cell detection logic
            return (1, 1);
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

    public class GridCell
    {
        /// <summary>
        /// The underlying Word.Cell this grid cell wraps.
        /// </summary>
        public Word.Cell COMCell { get; private set; }

        /// <summary>
        /// True if this cell is the top-left (master) of a merged group.
        /// </summary>
        public bool IsMasterCell { get; private set; }

        /// <summary>
        /// Zero-based row index within the internal table grid.
        /// </summary>
        public int GridRow { get; private set; }

        /// <summary>
        /// Zero-based column index within the internal table grid.
        /// </summary>
        public int GridCol { get; private set; }

        public GridCell(Word.Cell cell, int gridRow, int gridCol, bool isMasterCell)
        {
            COMCell = cell;
            GridRow = gridRow;
            GridCol = gridCol;
            IsMasterCell = isMasterCell;
        }
    }
}
