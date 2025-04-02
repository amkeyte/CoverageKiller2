using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    public static class TableGridCrawler3
    {
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

            // 🔍 Analyze merge spans
            AnalyzeSpans(jagged);

            return jagged;
        }

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
                        //throw new Exception("something wrong");
                        // If COM freaks out, keep going
                    }
                }
            }
        }

        public static void PrepareGridForLayout(Base1JaggedList<GridCell2> grid, float rowHeight = 20f, float colWidth = 20f)
        {
            try
            {
                var table = grid.Get2D(1, 1).COMCell.Range.Tables[1];
                //table.AllowAutoFit = false;

                table.Range.Font.Name = "Consolas";
                table.Range.Font.Size = 10;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = grid[1].Count * colWidth;

                foreach (var row in grid)
                {
                    Debug.WriteLine(row[1].COMCell.Width);
                }
            }
            //foreach (var row in grid)
            //{
            //    foreach (var cell in row)
            //    {
            //        if (cell.IsDummy || cell.COMCell == null)
            //            continue;

            //        try
            //        {
            //            var com = cell.COMCell;

            //            com.Row.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            //            com.Row.Height = rowHeight;

            //            com.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            //            com.PreferredWidth = colWidth;

            //            com.Range.Font.Name = "Consolas";
            //            com.Range.Font.Size = 10;
            //            com.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //        }
            catch
            {
                if (Debugger.IsAttached) Debugger.Break();
                // Word will sometimes throw if a merged ghost is touched — skip them safely
            }
        }


        //var firstReal = grid.SelectMany(r => r).FirstOrDefault(c => !c.IsDummy && c.COMCell != null);
        //firstReal?.COMCell?.Range?.ParagraphFormat?.SetSpaceAfter(0);


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

                // Horizontal span
                int colSpan = 1;
                for (int c = col + 1; c <= colCount; c++)
                {
                    var right = grid[row][c];
                    if (right is ZombieCell2 z && z.MasterCell == cell)
                        colSpan++;
                    else
                        break;
                }

                // Vertical span
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

    public class GridCell2
    {
        private RangeSnapshot snapshot;

        public Word.Cell COMCell { get; protected set; }
        public int GridRow { get; protected set; }
        public int GridCol { get; protected set; }
        public bool IsMasterCell { get; protected set; }
        public GridCell2 MasterCell { get; protected set; }

        public int DetectedRowSpan { get; set; } = 1;
        public int DetectedColSpan { get; set; } = 1;

        public RangeSnapshot Snapshot
        {
            get
            {
                snapshot = COMCell is null ? null : new RangeSnapshot(COMCell.Range);
                return snapshot;
            }
            protected set => snapshot = value;
        }

        public GridCell2(Word.Cell cell, int gridRow, int gridCol, bool isMasterCell)
        {
            COMCell = cell;
            GridRow = gridRow;
            GridCol = gridCol;
            IsMasterCell = isMasterCell;
            MasterCell = this;
        }

        public bool IsMerged => MasterCell != this;
        public virtual bool IsDummy => false;
    }

    public class ZombieCell2 : GridCell2
    {
        public ZombieCell2(GridCell2 master, int gridRow, int gridCol)
            : base(null, gridRow, gridCol, false)
        {
            COMCell = null;
            MasterCell = master;
        }

        public override bool IsDummy => true;
    }
}
