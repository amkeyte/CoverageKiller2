using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a grid abstraction for a Word table,
    /// mapping the table cells into a 2D array using X/Y coordinates.
    /// </summary>
    public class CKTableGrid
    {
        public static CKTableGrid GetInstance(Word.Table table)
        {
            var tableRange = new CKRange(table.Range);

            if (_tableGrids.TryGetValue(tableRange, out CKTableGrid grid))
            {
                return grid;
            }
            else
            {
                grid = new CKTableGrid(table);
                _tableGrids.Add(tableRange, grid);
                return grid;
            }
        }

        public static IEnumerable<GridChange> Refresh(CKTable table, bool getDiffs = false)
        {
            var oldGrid = _tableGrids[table];
            var newGrid = new CKTableGrid(table.COMTable);
            _tableGrids[table] = new CKTableGrid(table.COMTable);

            if (getDiffs)
            {
                return CKTableGridDiff.DiffGrids(oldGrid, newGrid);
            }

            return null;
        }

        private static Dictionary<CKRange, CKTableGrid> _tableGrids = new Dictionary<CKRange, CKTableGrid>();

        private Word.Table _table;
        private GridCell[,] _grid;

        /// <summary>
        /// Gets the number of rows in the grid (0-based size).
        /// </summary>
        public int RowCount { get; internal set; }
        /// <summary>
        /// Gets the number of columns in the grid (0-based size).
        /// </summary>
        public int ColCount { get; internal set; }

        /// <summary>
        /// Constructs a grid from the given Word table.
        /// </summary>
        /// <param name="table">The Word table to map.</param>
        private CKTableGrid(Word.Table table)
        {
            _table = table;
            BuildGrid();
        }

        /// <summary>
        /// Retrieves the grid cell at the given row and column indices (0-based).
        /// Returns null if the coordinates are out of range.
        /// </summary>
        public GridCell GetCellAt(int row, int col)
        {
            if (row < 0 || row >= RowCount ||
                col < 0 || col >= ColCount)
                return null;
            return _grid[row, col];
        }

        public GridCell GetMasterGridCellForWordCell(CKCell cell)
        {
            return _grid.Cast<GridCell>()
                        .Where(gc => gc != null && gc.COMCell == cell.COMCell && gc.IsMasterCell)
                        .FirstOrDefault();
        }

        /// <summary>
        /// Builds the grid representation from the Word table.
        /// </summary>
        private void BuildGrid()
        {
            Word.Table wordTable = _table;
            int rowCount = wordTable.Rows.Count;
            int colCount = GetMaxColumns(wordTable);

            // Set grid dimensions.
            RowCount = rowCount;
            ColCount = colCount;

            // Initialize the grid (using 0-based indexing)
            _grid = new GridCell[rowCount, colCount];

            // Iterate over each row in the table
            for (int i = 1; i <= rowCount; i++)  // Word Interop collections are 1-based
            {
                Word.Row row = wordTable.Rows[i];
                int currentGridCol = 0;

                // Iterate over each cell in the current row
                for (int j = 1; j <= row.Cells.Count; j++)
                {
                    Word.Cell cell = row.Cells[j];

                    // Advance currentGridCol past any already filled grid positions.
                    while (currentGridCol < colCount && _grid[i - 1, currentGridCol] != null)
                    {
                        currentGridCol++;
                    }

                    if (currentGridCol >= colCount)
                        break; // Safety check

                    // Get the cell span (rowSpan, colSpan) for merged cells.
                    // For now, this is a stub returning (1,1). Enhance it as needed.
                    (int rowSpan, int colSpan) = GetCellSpan(cell);

                    // Populate the grid for the cell. The master cell (top-left) gets marked.
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
        /// <summary>
        /// Determines the maximum number of columns in the table.
        /// For a table without merged cells, this is the maximum Cells.Count across all rows.
        /// </summary>
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

        /// <summary>
        /// Stub implementation that returns the span for a given cell.
        /// In a real implementation, you would determine if the cell is merged horizontally
        /// and/or vertically, and return the appropriate rowSpan and colSpan.
        /// </summary>
        private (int rowSpan, int colSpan) GetCellSpan(Word.Cell cell)
        {
            // TODO: Implement logic to detect merged cells.
            // This might involve comparing the cell's width to the expected width,
            // or examining the underlying XML. For now, assume no merging.
            return (1, 1);
        }


        public CKRow GetRowCells(int rowNumber)
        {
            // Convert the 1-based row number (Word Interop) to 0-based index for the grid.
            int gridRowIndex = rowNumber - 1;
            int colCount = _grid.GetLength(1);

            // For each column in the specified row, select non-null GridCells.
            // Then group them by their underlying Word.Cell (so merged cells appear only once),
            // order by the first occurrence (lowest column index), and return the Word.Cell.
            var rowCells = Enumerable.Range(0, colCount)
                .Select(col => new { ColIndex = col, GridCell = _grid[gridRowIndex, col] })
                .Where(x => x.GridCell != null)
                .GroupBy(x => x.GridCell.COMCell)
                .Select(g => new { FirstCol = g.Min(x => x.ColIndex), GridCell = g.First().GridCell })
                .OrderBy(x => x.FirstCol)
                .Select(x => x.GridCell);

            return new CKRow(rowCells);
        }

        public CKColumn GetColumnCells(int columnNumber)
        {
            // Convert the 1-based column number (Word Interop) to a 0-based index for our grid.
            int gridColIndex = columnNumber - 1;
            int rowCount = _grid.GetLength(0);

            var columnGridCells = Enumerable.Range(0, rowCount)
                .Select(row => new { RowIndex = row, GridCell = _grid[row, gridColIndex] })
                .Where(x => x.GridCell != null)
                .GroupBy(x => x.GridCell.COMCell)
                .Select(g => new { FirstRow = g.Min(x => x.RowIndex), GridCell = g.First().GridCell })
                .OrderBy(x => x.FirstRow)
                .Select(x => x.GridCell);



            return new CKColumn(null);
        }
    }


    /// <summary>
    /// Represents a cell in the grid, wrapping a Word.Cell.
    /// The IsMasterCell flag indicates if this cell is the top-left cell of a merged group.
    /// </summary>
    public class GridCell
    {
        public Word.Cell COMCell { get; private set; }
        public CKCell Item => new CKCell(COMCell);
        public bool IsMasterCell { get; private set; }
        public int Row { get; private set; }
        public int Col { get; private set; }

        public GridCell(Word.Cell cell, int row, int col, bool isMasterCell)
        {
            COMCell = cell;
            Row = row;
            Col = col;
            IsMasterCell = isMasterCell;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (!(obj is GridCell))
                return false;
            GridCell other = (GridCell)obj;
            return IsMasterCell == other.IsMasterCell &&
                   COMCell == other.COMCell &&
                   Row == other.Row &&
                   Col == other.Col;
        }

        public override int GetHashCode()
        {
            // Generate a hash code based on the key properties.
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + IsMasterCell.GetHashCode();
                hash = hash * 23 + (COMCell?.GetHashCode() ?? 0);
                hash = hash * 23 + Row;
                hash = hash * 23 + Col;
                return hash;
            }
        }
    }

    public static class CKTableGridDiff
    {
        /// <summary>
        /// Diffs two grid snapshots and returns a list of individual grid changes.
        /// </summary>
        public static List<GridChange> DiffGrids(CKTableGrid oldGrid, CKTableGrid newGrid)
        {
            var changes = new List<GridChange>();

            int oldRowCount = oldGrid.RowCount;
            int oldColCount = oldGrid.ColCount;
            int newRowCount = newGrid.RowCount;
            int newColCount = newGrid.ColCount;

            // Detect row count changes.
            if (newRowCount != oldRowCount)
            {
                // For simplicity, we flag a single change at the first differing row.
                int diffIndex = FindFirstRowDifference(oldGrid, newGrid, Math.Min(oldRowCount, newRowCount));
                if (newRowCount > oldRowCount)
                {
                    changes.Add(new GridChange { ChangeType = GridChangeType.RowInserted, RowIndex = diffIndex });
                }
                else
                {
                    changes.Add(new GridChange { ChangeType = GridChangeType.RowDeleted, RowIndex = diffIndex });
                }
            }

            // Detect column count changes.
            if (newColCount != oldColCount)
            {
                int diffIndex = FindFirstColumnDifference(oldGrid, newGrid, Math.Min(oldColCount, newColCount));
                if (newColCount > oldColCount)
                {
                    changes.Add(new GridChange { ChangeType = GridChangeType.ColumnInserted, ColumnIndex = diffIndex });
                }
                else
                {
                    changes.Add(new GridChange { ChangeType = GridChangeType.ColumnDeleted, ColumnIndex = diffIndex });
                }
            }

            // Compare overlapping cells.
            int minRows = Math.Min(oldRowCount, newRowCount);
            int minCols = Math.Min(oldColCount, newColCount);
            for (int r = 0; r < minRows; r++)
            {
                for (int c = 0; c < minCols; c++)
                {
                    GridCell oldCell = oldGrid.GetCellAt(r, c);
                    GridCell newCell = newGrid.GetCellAt(r, c);
                    if (!Equals(oldCell, newCell))
                    {
                        changes.Add(new GridChange
                        {
                            ChangeType = GridChangeType.CellModified,
                            RowIndex = r,
                            ColumnIndex = c,
                            OldCell = oldCell,
                            NewCell = newCell
                        });
                    }
                }
            }

            return changes;
        }

        /// <summary>
        /// Finds the first row index where the two grids differ.
        /// </summary>
        private static int FindFirstRowDifference(CKTableGrid oldGrid, CKTableGrid newGrid, int rowCount)
        {
            int minCols = Math.Min(oldGrid.ColCount, newGrid.ColCount);
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < minCols; c++)
                {
                    GridCell oldCell = oldGrid.GetCellAt(r, c);
                    GridCell newCell = newGrid.GetCellAt(r, c);
                    if (!Equals(oldCell, newCell))
                        return r;
                }
            }
            return 0;
        }

        /// <summary>
        /// Finds the first column index where the two grids differ.
        /// </summary>
        private static int FindFirstColumnDifference(CKTableGrid oldGrid, CKTableGrid newGrid, int colCount)
        {
            int minRows = Math.Min(oldGrid.RowCount, newGrid.RowCount);
            for (int c = 0; c < colCount; c++)
            {
                for (int r = 0; r < minRows; r++)
                {
                    GridCell oldCell = oldGrid.GetCellAt(r, c);
                    GridCell newCell = newGrid.GetCellAt(r, c);
                    if (!Equals(oldCell, newCell))
                        return c;
                }
            }
            return 0;
        }
    }
    public enum GridChangeType
    {
        RowInserted,
        RowDeleted,
        ColumnInserted,
        ColumnDeleted,
        CellModified  // e.g. a cell's properties changed (position, master cell, or COMCell reference)
    }

    public class GridChange
    {
        public GridChangeType ChangeType { get; set; }

        // 0-based indices for rows and columns where the change occurred.
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }

        // For cell modifications, include the before/after state.
        public GridCell OldCell { get; set; }
        public GridCell NewCell { get; set; }

        // Optionally, add additional properties if needed (e.g., number of rows/columns inserted/deleted).
    }


}
