using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Utility class for visualizing and processing Word table grid layouts using GridCell2 structures.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0013
    /// </remarks>
    public static class TableGridCrawler3
    {
        /// <summary>
        /// Dumps a string representation of the grid for debugging or visualization.
        /// </summary>
        /// <param name="grid">The jagged grid structure of GridCell2 instances.</param>
        /// <returns>A formatted string showing grid positions and cell types.</returns>
        public static string DumpGrid(Base1JaggedList<GridCell2> grid)
        {
            var sb = new StringBuilder();
            foreach (var row in grid)
            {
                var line = row.Select(c =>
                {
                    var label = c.IsDummy ? "Z" : (c.IsMerged ? "M" : "O");
                    return $"{label}[{c.GridRow},{c.GridCol}]";
                });

                sb.AppendLine(string.Join(" | ", line));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Constructs a normalized grid layout of the Word table using GridCell2 and ZombieCell2 placeholders.
        /// Pads gaps and analyzes spans due to merged cells.
        /// </summary>
        /// <param name="table">The Word table to process.</param>
        /// <returns>A padded jagged grid representation of the table.</returns>
        public static Base1JaggedList<GridCell2> NormalizeVisualGrid(Word.Table table)
        {
            var raw = table.Range.Cells
                .Cast<Word.Cell>()
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(c => c.ColumnIndex).ToList())
                .ToList();

            var jagged = new Base1JaggedList<GridCell2>();
            int globalMaxCols = 0;

            foreach (var wordRow in raw)
            {
                var row = new Base1List<GridCell2>();
                int currentCol = 1;
                GridCell2 lastRealCell = null;

                foreach (var cell in wordRow)
                {
                    int gap = cell.ColumnIndex - currentCol;

                    // Fill gap with dummy placeholders
                    for (int i = 0; i < gap; i++)
                    {
                        var dummy = new ZombieCell2(lastRealCell, wordRow[0].RowIndex, currentCol);
                        row.Add(dummy);
                        currentCol++;
                    }

                    var real = new GridCell2(cell, cell.RowIndex, cell.ColumnIndex, true);
                    row.Add(real);
                    lastRealCell = real;
                    currentCol++;
                }

                globalMaxCols = Math.Max(globalMaxCols, row.Count);
                jagged.Add(row);
            }

            // Vertically pad short rows
            foreach (var row in jagged)
            {
                while (row.Count < globalMaxCols)
                {
                    var master = row.FirstOrDefault(c => !c.IsDummy);
                    int col = row.Count + 1;
                    int rowIndex = master?.GridRow ?? jagged.IndexOf(row);
                    var dummy = new ZombieCell2(master, rowIndex, col);
                    row.Add(dummy);
                }
            }

            AnalyzeSpans(jagged);
            return jagged;
        }

        /// <summary>
        /// Colors all non-master cells based on their master cell's background.
        /// Used for visual debugging.
        /// </summary>
        /// <param name="grid">The grid to process.</param>
        public static void ColorMasterCells(Base1JaggedList<GridCell2> grid)
        {
            foreach (var cell in grid.SelectMany(r => r))
            {
                if (!cell.IsMasterCell)
                {
                    try
                    {
                        cell.MasterCell.COMCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightBlue;
                    }
                    catch
                    {
                        // Skip safely if COM is unavailable or misbehaving
                    }
                }
            }
        }

        /// <summary>
        /// Prepares the table for layout rendering by adjusting width, font, and formatting.
        /// </summary>
        /// <param name="grid">The normalized grid.</param>
        /// <param name="rowHeight">Optional row height in points.</param>
        /// <param name="colWidth">Optional column width in points.</param>
        public static void PrepareGridForLayout(Base1JaggedList<GridCell2> grid, float rowHeight = 20f, float colWidth = 20f)
        {
            try
            {
                var table = grid.Get2D(1, 1).COMCell.Range.Tables[1];

                table.Range.Font.Name = "Consolas";
                table.Range.Font.Size = 10;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = grid[1].Count * colWidth;

                foreach (var row in grid)
                {
                    Debug.WriteLine(row[1].COMCell.Width);
                }
            }
            catch
            {
                if (Debugger.IsAttached) Debugger.Break();
            }
        }

        /// <summary>
        /// Detects row and column span for each master cell based on attached dummy neighbors.
        /// </summary>
        /// <param name="grid">The normalized grid to process.</param>
        private static void AnalyzeSpans(Base1JaggedList<GridCell2> grid)
        {
            int rowCount = grid.Count;
            int colCount = grid[1].Count;

            foreach (var cell in grid.SelectMany(row => row))
            {
                if (cell.IsDummy || !cell.IsMasterCell || cell.COMCell == null)
                    continue;

                int row = cell.GridRow;
                int col = cell.GridCol;

                int colSpan = 1;
                for (int c = col + 1; c <= colCount; c++)
                {
                    var right = grid[row][c];
                    if (right is ZombieCell2 z && z.MasterCell == cell)
                        colSpan++;
                    else
                        break;
                }

                int rowSpan = 1;
                for (int r = row + 1; r <= rowCount; r++)
                {
                    var down = grid[r][col];
                    if (down is ZombieCell2 z && z.MasterCell == cell)
                        rowSpan++;
                    else
                        break;
                }

                cell.DetectedColSpan = colSpan;
                cell.DetectedRowSpan = rowSpan;
            }
        }
    }

    /// <summary>
    /// Represents a cell in a Word table, as a coordinate-aware grid object.
    /// Tracks merge structure and provides reference to the COM cell.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0013
    /// </remarks>
    public class GridCell2
    {
        private RangeSnapshot snapshot;

        /// <summary>
        /// The underlying Word.Cell COM object (null if dummy).
        /// </summary>
        public Word.Cell COMCell { get; protected set; }

        /// <summary>
        /// One-based grid row coordinate.
        /// </summary>
        public int GridRow { get; protected set; }

        /// <summary>
        /// One-based grid column coordinate.
        /// </summary>
        public int GridCol { get; protected set; }

        /// <summary>
        /// Indicates if this is the master cell of a merged region.
        /// </summary>
        public bool IsMasterCell { get; protected set; }

        /// <summary>
        /// The associated master cell for this position (self if not merged).
        /// </summary>
        public GridCell2 MasterCell { get; protected set; }

        /// <summary>
        /// Row span detected by `AnalyzeSpans`.
        /// </summary>
        public int DetectedRowSpan { get; set; } = 1;

        /// <summary>
        /// Column span detected by `AnalyzeSpans`.
        /// </summary>
        public int DetectedColSpan { get; set; } = 1;

        /// <summary>
        /// Snapshot of the cell’s text range and metadata.
        /// </summary>
        public RangeSnapshot Snapshot
        {
            get
            {
                snapshot = COMCell is null ? null : new RangeSnapshot(COMCell.Range);
                return snapshot;
            }
            protected set => snapshot = value;
        }

        /// <summary>
        /// Constructs a new grid cell for a real (non-dummy) Word.Cell.
        /// </summary>
        public GridCell2(Word.Cell cell, int gridRow, int gridCol, bool isMasterCell)
        {
            COMCell = cell;
            GridRow = gridRow;
            GridCol = gridCol;
            IsMasterCell = isMasterCell;
            MasterCell = this;
        }

        /// <summary>
        /// Indicates if this cell is part of a merged group.
        /// </summary>
        public bool IsMerged => MasterCell != this;

        /// <summary>
        /// Indicates if this cell is a dummy placeholder.
        /// </summary>
        public virtual bool IsDummy => false;
    }

    /// <summary>
    /// Represents a dummy cell used to fill gaps in the visual grid due to merged cells or row padding.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0013
    /// </remarks>
    public class ZombieCell2 : GridCell2
    {
        /// <summary>
        /// Constructs a new dummy grid cell.
        /// </summary>
        /// <param name="master">The master cell that owns this dummy cell.</param>
        /// <param name="gridRow">The grid row index of the dummy.</param>
        /// <param name="gridCol">The grid column index of the dummy.</param>
        public ZombieCell2(GridCell2 master, int gridRow, int gridCol)
            : base(null, gridRow, gridCol, false)
        {
            COMCell = null;
            MasterCell = master;
        }

        /// <inheritdoc/>
        public override bool IsDummy => true;
    }
}
