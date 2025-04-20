using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Represents a visual layout cell with merge span metadata,
    /// suitable for use as a reusable layout template independent of COM references.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0033
    /// </remarks>
    public class GridCell5
    {
        public int GridRow { get; }
        public int GridCol { get; }
        public virtual bool IsRowEndMarker { get; } = false;
        public virtual bool IsMasterCell { get; } = false;
        public virtual GridCell5 MasterCell { get; }
        public virtual bool IsGhostCell { get; } = false;
        public virtual bool IsMergedCell { get; } = false;

        public int ColSpan { get; private set; } = 1;
        public int RowSpan { get; private set; } = 1;

        public GridCell5(int row, int col, GridCell5 masterCell = null)
        {
            GridRow = row;
            GridCol = col;
            MasterCell = masterCell ?? this;
            IsMasterCell = true;
        }

        internal void AddMerge(MergedGridCell5 mergedGridCell5)
        {
            //if a merged cell is tied to this one that is further than the current span, give it the new span.
            //add plus 1, because span is inclusive.
            ColSpan = Math.Max(ColSpan, (mergedGridCell5.GridCol - GridCol) + 1);
            RowSpan = Math.Max(RowSpan, (mergedGridCell5.GridRow - GridRow) + 1);
        }
    }
    /// <summary>
    /// Ghost cell points to itself but is not a master. used to take place
    /// of undiscovered grid placements.
    /// </summary>
    public class GhostGridCell5 : GridCell5
    {
        public override bool IsMasterCell => false;
        public override bool IsGhostCell => true;
        public override GridCell5 MasterCell => null;
        public GhostGridCell5()
            : base(-1, -1, null)
        {
        }

    }
    public class RowEndGridCell5 : GridCell5
    {
        public override bool IsMasterCell => false;
        public override GridCell5 MasterCell => null;
        public override bool IsRowEndMarker => true;
        public RowEndGridCell5()
            : base(-999, -1, null)
        {

        }
    }
    public class MergedGridCell5 : GridCell5
    {
        public override bool IsMasterCell => false;
        public override bool IsMergedCell => true;

        public MergedGridCell5(int row, int col, GridCell5 masterCell)
            : base(row, col, masterCell)
        {
            masterCell.AddMerge(this);
        }
    }


    /// <summary>
    /// Builds a reusable layout grid from a Word table using both visual and textual merge inference.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0034
    /// </remarks>
    public class GridCrawler5
    {
        private readonly CKTable _table;
        private Base1JaggedList<GridCell5> _grid;
        private Base1JaggedList<GridCell5> _masterCells;

        /// <summary>
        /// Initializes and analyzes the layout grid for the specified table.
        /// </summary>
        /// <param name="table">The source table to analyze.</param>
        public GridCrawler5(CKTable table)
        {
            this.Ping();
            _table = table ?? throw new ArgumentNullException(nameof(table));
            Analyze();
            this.Pong();
        }

        /// <summary>
        /// Gets the number of visual rows in the final layout grid.
        /// </summary>
        public int RowCount => _grid.Count;

        /// <summary>
        /// Gets the number of visual columns in the final layout grid.
        /// </summary>
        public int ColumnCount => _grid.LargestRowCount;

        /// <summary>
        /// Gets the completed visual layout grid.
        /// </summary>
        public Base1JaggedList<GridCell5> Grid => _grid;

        /// <summary>
        /// Returns a labeled text representation of the grid.
        /// </summary>
        public static string DumpGrid(Base1JaggedList<GridCell5> grid, string message = null)
        {
            var sb = new StringBuilder();
            sb = sb.AppendLine(message + "\n");
            sb.AppendLine("**********************");

            foreach (var row in grid)
            {
                string line = default;

                foreach (var cell in row)
                {
                    line = "[ERROR]";
                    if (cell.IsMasterCell)
                    {
                        line = $"M[{cell.GridRow},{cell.GridCol}]";
                    }
                    else if (cell.IsMergedCell)
                    {
                        line = $"*[{cell.MasterCell.GridRow},{cell.MasterCell.GridCol}]";
                    }
                    else if (cell.IsRowEndMarker)
                    {
                        line = $"[END]\n";
                    }
                    else if (cell.IsGhostCell)
                    {
                        line = $"[???]";
                    }
                    else
                    {
                        throw new Exception("wtf?");
                    }

                    sb.Append(line + " | ");
                }
            }
            sb.AppendLine("**********************");

            return sb.ToString();
        }
        /// <summary>
        /// Returns a labeled text representation of the grid.
        /// </summary>
        public static string DumpGrid(Base1JaggedList<string> grid, string message = null)
        {
            var sb = new StringBuilder();
            sb = sb.AppendLine(message + "\n");
            sb.AppendLine("**********************");

            foreach (var row in grid)
            {
                var line = row.Select(cell => string.IsNullOrEmpty(cell) ? "[NULL]" : cell);
                sb.AppendLine(string.Join(" | ", line));
            }
            sb.AppendLine("**********************");

            return sb.ToString();
        }
        /// <summary>
        /// Dumps a string representation of the grid for debugging or visualization.
        /// </summary>
        /// <param name="grid">The jagged grid structure of GridCell2 instances.</param>
        /// <returns>A formatted string showing grid positions and cell types.</returns>
        public static string DumpGrid(Base1JaggedList<Word.Cell> grid, string message = null)
        {
            var sb = new StringBuilder();
            sb = sb.AppendLine(message + "\n");
            sb.AppendLine("**********************");

            foreach (var row in grid)
            {
                var line = row.Select(c =>
                {

                    return $"[{c.RowIndex},{c.ColumnIndex}]";
                });

                sb.AppendLine(string.Join(" | ", line));
            }
            sb.AppendLine("**********************");
            return sb.ToString();
        }
        // ========== Internal Analysis Flow ==========

        private void Analyze()
        {
            LH.Ping(GetType());

            LH.Checkpoint($"Cloning table {_table.Document.Tables.IndexOf(_table)} from document {_table.Document.FileName}");
            var clonedTable = CloneAndPrepareTableLayout();
            var masterGrid = GetMasterGrid(clonedTable);
            var textGrid = ParseTableText(clonedTable);
            var normalGrid = NormalizeByWidth(masterGrid);
            var horizGrid = CrawlHoriz(textGrid, normalGrid);
            CrawlVertically(textGrid, normalGrid);

            LH.Pong(GetType());
        }

        public static Base1List<string> SplitWordTableTextIntoRows(string rawText)
        {

            var rows = new Base1List<string>();
            if (string.IsNullOrEmpty(rawText)) return rows;

            int current = 0;
            while (current < rawText.Length)
            {
                int nextBreak = rawText.IndexOf("\r\a\r\a", current, StringComparison.Ordinal);
                if (nextBreak < 0)
                {
                    rows.Add(rawText.Substring(current));
                    break;
                }

                int rowEnd = nextBreak + 4; // includes both \r\a\r\a
                string rowText = rawText.Substring(current, rowEnd - current);
                rows.Add(rowText);
                current = rowEnd;
            }

            return rows;
        }




        /// <summary>
        /// Parses the table's raw emitted text into a jagged list of cell strings.
        /// Accounts for Word's \r\a patterns and merged cell behavior.
        /// </summary>
        /// <returns>A jagged list of strings representing logical cell values.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0037
        /// </remarks>
        internal Base1JaggedList<string> ParseTableText(CKTable table = null)
        {
            LH.Ping(GetType());



            table = table ?? _table;
            //string rawText = table.Text; // Use the CKTable abstraction
            //var parts = rawText.Split(new[] { "\r\a" }, StringSplitOptions.None).ToList();
            var colCount = GetMasterGrid(table).LargestRowCount;



            var result = new Base1JaggedList<string>();
            //var currentRow = new Base1List<string>();

            var rowTexts = SplitWordTableTextIntoRows(table.RawText);



            Log.Verbose(DumpList(rowTexts,
                $"Parsed Table[{table.Document.Tables.IndexOf(table)}] Text"));


            foreach (var row in rowTexts)
            {
                //var parts = row.Split(new[] { "\r\a" }, StringSplitOptions.None).ToList();
                var matches = Regex.Matches(row, @"(.*?\r\a)")
                 .Cast<Match>()
                 .Select(m => m.Value)
                 .ToList();

                var parts = new List<string>();
                foreach (var part in matches)
                {
                    parts.Add(FlattenTableText(part));
                }

                //parts.RemoveAt(parts.Count - 1); //remove last because the double tap makes it add an "extra"
                result.Add(new Base1List<string>(parts));
            }
            for (int i = 1; i < result.Count; i++)
            {
                //shuffle the nulls forward in the list

                var row = result[i];
                var nextRowIndex = i + 1;
                while (row.Count < colCount + 1 && result.Count >= nextRowIndex)
                {
                    row.Add(result[nextRowIndex][1]);
                    result[nextRowIndex].RemoveAt(1);
                    if (result[nextRowIndex].Count == 0) result.RemoveAt(nextRowIndex);//maybe not save while in row iteration.
                }

            }

            //result.Where(r => !r.Any())
            //    .Reverse().ToList()
            //    .ForEach(r => result.RemoveAt(result.IndexOf(r)));

            LH.Pong(GetType());

            _textGrid = result;
            return result;
        }

        /// <summary>
        /// Dumps the raw contents of a Base1List&lt;string&gt;, each entry as-is, separated by newlines.
        /// </summary>
        /// <param name="text">The list to dump.</param>
        /// <param name="v">Unused label (included for signature compatibility).</param>
        /// <returns>The concatenated raw text of the list.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0036
        /// </remarks>
        public static string DumpList(Base1List<string> textList, string message)
        {
            var sb = new StringBuilder();
            sb = sb.AppendLine("\n" + message + "\n");
            sb.AppendLine("**********************");

            foreach (var row in textList)
            {
                sb.AppendLine(FlattenTableText(row));
            }
            sb.AppendLine("**********************");

            return sb.ToString();
        }


        public static string FlattenTableText(string tableText)
        {

            return tableText
                .Replace("\r", "/r")
                .Replace("\a", "/a")
                .Replace("\n", "/n")
                .Replace("\\", "/");
        }
        /// <summary>
        /// Constructs the GridCell layout with dummy cells for visual alignment.
        /// </summary>
        internal Base1JaggedList<GridCell5> NormalizeByWidth(
            Base1JaggedList<Word.Cell> masterGrid = null)
        {
            LH.Ping(GetType());

            var result = new Base1JaggedList<GridCell5>();
            // Step 1: Build master-only grid
            masterGrid = masterGrid ?? GetMasterGrid(_table);

            if (masterGrid.Count == 0)
                throw new InvalidOperationException("Table has no master grid rows.");

            // Step 2: Find the row with the most master cells
            var widestRow = masterGrid.OrderByDescending(r => r.Count).First();
            float totalRowWidth = widestRow.Sum(c => c.Width);
            int colCount = widestRow.Count;

            // Step 3: Average width per column (baseline)
            float normalWidth = totalRowWidth / colCount;


            // Step 4: Pad each row by inserting ZombieCells after wide master cells
            var newGrid = new Base1JaggedList<GridCell5>();

            foreach (var row in masterGrid)
            {
                var newRow = new Base1List<GridCell5>();
                //insert cells where there are wide spaces
                foreach (var cell in row)
                {
                    var newCell = new GridCell5(cell.RowIndex, cell.ColumnIndex);
                    newRow.Add(newCell);

                    int span = Math.Max(1, (int)Math.Round(cell.Width / normalWidth));

                    for (int i = 1; i < span; i++)
                    {
                        newRow.Add(new MergedGridCell5(cell.RowIndex, cell.ColumnIndex + i, newCell));
                    }

                }
                //insert cells to fill out row
                for (int i = newRow.Count; i < colCount; i++)
                {
                    newRow.Add(new GhostGridCell5());
                }
                newRow.Add(new RowEndGridCell5());
                newGrid.Add(newRow);
            }
            Log.Debug(GridCrawler5.DumpGrid(newGrid));

            _grid = newGrid;
            LH.Pong(GetType());

            return _grid;
        }
        /// <summary>
        /// Clones the given Word table into a shadow workspace and formats it for grid-based layout visualization.
        /// </summary>
        /// <param name="sourceTable">The table to clone and prepare.</param>
        /// <param name="colWidth">Optional column width in points (default: 20f).</param>
        /// <returns>A ShadowWorkspace containing the formatted cloned table and master cell grid.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0015
        /// </remarks>
        public CKTable CloneAndPrepareTableLayout(
            CKTable sourceTable = null,
            ShadowWorkspace workspace = null)
        {
            LH.Ping(GetType());

            sourceTable = sourceTable ?? _table;
            workspace = workspace ?? sourceTable.Application.GetShadowWorkspace();

            //for debugging uncomment.
            workspace.ShowDebuggerWindow();

            //put original table
            workspace.CloneFrom(_table);
            workspace.Document.Content.CollapseToEnd().Text = "\r\r\r";
            //put the one to format
            var clonedTable = workspace.CloneFrom(sourceTable);
            //var grid = GetMasterGrid(clonedTable);
            //Log.Debug(GridCrawler5.DumpGrid(grid));


            //normalize the text
            clonedTable.COMRange.Font.Name = "Consolas";
            clonedTable.COMRange.Font.Size = 10;

            // Enable auto-fit so columns stretch evenly to fill the table width
            clonedTable.COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            // set the table width to make.
            clonedTable.COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            clonedTable.COMTable.PreferredWidth = 100f;

            int cellCounter = 1;
            foreach (Word.Cell cell in clonedTable.COMRange.Cells)
            {
                cell.Range.Text = cellCounter++.ToString();
            }

            LH.Pong(GetType());

            return clonedTable;
        }
        /// <summary>
        /// Retrieves a grid of master GridCell5s from the CKTable.
        /// </summary>
        public Base1JaggedList<Word.Cell> GetMasterGrid(CKTable table = null)
        {
            LH.Ping(GetType());

            table = table ?? _table;

            var groupedRows = table.COMTable.Range.Cells
                .Cast<Word.Cell>()
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(c => c.ColumnIndex))
                .ToList();

            var result = new Base1JaggedList<Word.Cell>();

            foreach (var row in groupedRows)
            {
                var list = new Base1List<Word.Cell>();
                foreach (var cell in row)
                {
                    list.Add(cell);
                }
                result.Add(list);
            }
            Log.Debug(GridCrawler5.DumpGrid(result));

            LH.Pong(GetType());

            return result;
        }

        Base1JaggedList<string> _textGrid = new Base1JaggedList<string>();
        internal Base1JaggedList<GridCell5> CrawlVertically(
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {
            LH.Ping(GetType());

            textGrid = textGrid ?? _textGrid;
            normalizedGrid = normalizedGrid ?? _grid;

            int rowCount = normalizedGrid.Count;

            int colCount = normalizedGrid[1].Count;

            for (int colIndex = 1; colIndex <= colCount; colIndex++)
            {

                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    var gridRow = normalizedGrid[rowIndex];
                    var gridCell = gridRow[colIndex];
                    var textCell = textGrid[rowIndex][colIndex];

                    if (gridCell.IsMasterCell)
                    {
                        if (textCell != "/r/a") continue;
                        if (Debugger.IsAttached) Debugger.Break();
                    }
                    else if (gridCell.IsMergedCell)
                    {
                        if (textCell == "/r/a") continue;
                        if (Debugger.IsAttached) Debugger.Break();
                    }
                    else if (gridCell.IsGhostCell)
                    {
                        if (textCell != "/r/a")
                        {
                            throw new Exception("Shouldn't get here.");
                        }

                        int up1Index = rowIndex - 1;
                        if (up1Index < 1) throw new Exception("can't go up");

                        var up1Row = normalizedGrid[up1Index];
                        var up1Cell = up1Row[colIndex];

                        if (up1Cell.IsMasterCell)
                        {
                            gridRow.RemoveAt(colIndex);
                            //if mastercell has a col span, insert enough cells to cover the merge rectangle.
                            for (var i = 0; i < up1Cell.ColSpan; i++)
                            {
                                gridRow.Insert(colIndex, new MergedGridCell5(up1Cell.GridRow, up1Cell.GridCol, up1Cell));
                            }
                            //gridCell = null;
                        }
                        else if (up1Cell.IsMergedCell)
                        {
                            gridRow.RemoveAt(colIndex);
                            for (var i = 0; i < up1Cell.MasterCell.ColSpan; i++)
                            {
                                gridRow.Insert(colIndex,
                                new MergedGridCell5(up1Cell.MasterCell.GridRow, up1Cell.MasterCell.GridCol, up1Cell.MasterCell));
                            }
                        }
                        else if (up1Cell.IsGhostCell)
                        {
                            throw new Exception("Shouldn't get here.");
                        }
                        else if (up1Cell.IsRowEndMarker)
                        {
                            //this was crap that got pushed over during a rectangular merge
                            gridRow.RemoveAt(colIndex);

                            //throw new Exception("Shouldn't get here either.");
                        }

                    }
                    else if (gridCell.IsRowEndMarker)
                    {
                        continue;
                    }
                }
            }
            _grid = normalizedGrid;



            Log.Debug(GridCrawler5.DumpGrid(textGrid));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid));

            LH.Pong(GetType());

            return normalizedGrid;
        }
        internal Base1JaggedList<GridCell5> CrawlHoriz(
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {
            LH.Ping(GetType());



            textGrid = textGrid ?? _textGrid;
            normalizedGrid = normalizedGrid ?? _grid;

            Log.Debug(DumpGrid(textGrid));
            Log.Debug(DumpGrid(normalizedGrid));


            var normalizedRowCount = normalizedGrid.LargestRowCount;

            for (var rowIndex = 1; rowIndex <= normalizedGrid.Count; rowIndex++)
            {
                var gridRow = normalizedGrid[rowIndex];
                var textRow = textGrid[rowIndex]; //figure out what to do if no next text row

                for (var cellIndex = 1; cellIndex <= normalizedRowCount; cellIndex++)
                {
                    var gridCell = gridRow[cellIndex];
                    var textCell = textRow[cellIndex];

                    if (gridCell.IsMasterCell)
                    {
                        if (cellIndex >= normalizedRowCount) throw new Exception(" master nope");

                        if (textCell != "/r/a") continue; //maybe compare actuaal text values?
                                                          //the text grid is correct.
                        gridRow.Insert(cellIndex, new GhostGridCell5());


                    }
                    else if (gridCell.IsMergedCell)
                    {
                        if (cellIndex >= normalizedRowCount) throw new Exception(" merged nope");

                        if (textCell == "/r/a") continue;

                        textRow.Insert(cellIndex, "/r/a");
                    }
                    else if (gridCell.IsRowEndMarker)
                    {

                        if (textCell != "/r/a") textRow.Insert(cellIndex, "/r/a");

                        for (var i = textRow.Count; i > cellIndex; i--)
                        {
                            //count from the back of the row and push them forward.
                            //add new row if needed

                            if (rowIndex + 1 > textGrid.Count) textGrid.Add(new Base1List<string>());
                            //insert the last cell from this row into the next one
                            textGrid[rowIndex + 1].Insert(1, textRow.Last());
                            //take the last cell off of this row
                            textRow.RemoveAt(textRow.Count);
                        }
                    }
                    else if (gridCell.IsGhostCell)
                    {
                        //cell is a shoved over from an erlier push
                        if (cellIndex >= normalizedRowCount)
                        {
                            //remove it and check it again
                            gridRow.RemoveAt(cellIndex);
                            cellIndex--;
                            continue;
                        }
                        //ghost cell is ok.
                        if (textCell == "/r/a") continue;

                        //musmatch, match text to ghost cell.
                        textRow.Insert(cellIndex, "/r/a");


                    }
                }

            }
            _textGrid = textGrid;
            _grid = normalizedGrid;

            Log.Debug(GridCrawler5.DumpGrid(textGrid));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid));

            LH.Pong(GetType());

            return normalizedGrid;
        }

    }
}
