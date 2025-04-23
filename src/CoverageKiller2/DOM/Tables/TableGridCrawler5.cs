using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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
    /// THIS CLASS IS NOT TO INCLUDE ANY CK.DOM references. Version: CK2.00.01.0033
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
    /// THIS CLASS IS NOT TO INCLUDE ANY CK.DOM references. Version: CK2.00.01.0034
    /// </remarks>
    public class GridCrawler5
    {
        private readonly Word.Table _COMTable;
        private Base1JaggedList<GridCell5> _grid;
        private Base1JaggedList<GridCell5> _masterCells;
        //private ShadowWorkspace _shadowWorkspace;
        internal GridCrawler5(Word.Table table)
        {
            this.Ping(msg: "$$$");
            _COMTable = table ?? throw new ArgumentNullException(nameof(table));
            //_shadowWorkspace = workspace ?? throw new ArgumentNullException(nameof(workspace));
            _grid = AnalyzeTableRecursively(table);
            this.Pong();
        }
        /// <summary>
        /// Initializes and analyzes the layout grid for the specified table.
        /// </summary>
        /// <param name="table">The source table to analyze.</param>
        //public GridCrawler5(CKTable table)
        //{
        //    throw new CKDebugException();
        //    //this.Ping(msg:"$$$");
        //    //Analyze();
        //    //this.Pong();
        //}

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
                        line = $"G[???]";
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

        //private void Analyze()
        //{
        //    this.Ping(msg: "$$$");

        //    //LH.Checkpoint($"Cloning table {_table.Document.Tables.IndexOf(_table)} from document {_table.Document.FileName}");

        //    //fast crawl

        //    //slow crawl
        //    var clonedTable = PrepareTable();
        //    var masterGrid = GetMasterGrid(clonedTable);
        //    var textGrid = ParseTableText(clonedTable);
        //    var normalGrid = NormalizeByWidth(masterGrid);
        //    var horizGrid = CrawlHoriz(textGrid, normalGrid);
        //    var vertGrid = CrawlVertically(0, textGrid, normalGrid);

        //    this.Pong();
        //}

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
        internal Base1JaggedList<string> ParseTableText(Word.Table COMTable = null)
        {
            this.Ping(msg: "$$$");



            COMTable = COMTable ?? _COMTable;
            var COMRange = COMTable.Range;
            //string rawText = table.Text; // Use the CKTable abstraction
            //var parts = rawText.Split(new[] { "\r\a" }, StringSplitOptions.None).ToList();
            var colCount = GetMasterGrid(COMTable).LargestRowCount;



            var result = new Base1JaggedList<string>();
            //var currentRow = new Base1List<string>();

            var rowTexts = SplitWordTableTextIntoRows(COMRange.Text);



            Log.Debug(DumpList(rowTexts, $"{nameof(ParseTableText)}\nRow Texts"));


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
            Log.Debug(DumpGrid(result, $"\n{nameof(ParseTableText)}-result"));

            this.Pong();

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
            Base1JaggedList<Word.Cell> masterGrid = null,
            int rowOffset = 0)
        {
            this.Ping(msg: "$$$");

            var result = new Base1JaggedList<GridCell5>();
            // Step 1: Build master-only grid
            masterGrid = masterGrid ?? GetMasterGrid(_COMTable);

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
                    var newCell = new GridCell5(cell.RowIndex + rowOffset, cell.ColumnIndex);
                    newRow.Add(newCell);

                    int span = Math.Max(1, (int)Math.Round(cell.Width / normalWidth));

                    for (int i = 1; i < span; i++)//for one cell, span is 1. 
                    {
                        newRow.Add(new MergedGridCell5(cell.RowIndex + rowOffset, cell.ColumnIndex + i, newCell));
                    }
                    span = 0;

                }
                //insert cells to fill out row
                for (int i = newRow.Count; i < colCount; i++)
                {
                    newRow.Add(new GhostGridCell5());
                }
                newRow.Add(new RowEndGridCell5());
                newGrid.Add(newRow);
            }
            Log.Debug(DumpGrid(newGrid, $"\n{nameof(NormalizeByWidth)}-{nameof(newGrid)}"));

            _grid = newGrid;
            this.Pong();

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
        public Word.Table PrepareTable(Word.Table COMTable = null)

        {
            this.Ping(msg: "$$$");

            COMTable = COMTable ?? _COMTable;
            var COMRange = _COMTable.Range;

            //normalize the text
            COMRange.Font.Name = "Consolas";
            COMRange.Font.Size = 10;

            // Enable auto-fit so columns stretch evenly to fill the table width
            COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            // set the table width to make.
            COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            COMTable.PreferredWidth = 100f;
            //get the longest row (assuming to have all the columns)
            var longestRow = GetMasterGrid(COMTable)
                .OrderByDescending(r => r.Count)
                .FirstOrDefault();
            //~78 X's will fit across the page. by cleaning out all content, then stuffing each cell full 
            //to word wrap, it should autosize each column an equal width.
            //this will help avoid a trap where autosze can be inconsistent causing phantom hits
            // on the width/span algorythm.
            string cellStretcher = new string('X', 100 / longestRow.Count);
            COMRange.Text = string.Empty;
            foreach (Word.Cell cell in longestRow)
            {
                cell.Range.Text = cellStretcher;
            }

            this.Pong();

            return COMTable;
        }
        /// <summary>
        /// Retrieves a grid of master GridCell5s from the CKTable.
        /// </summary>
        public Base1JaggedList<Word.Cell> GetMasterGrid(Word.Table COMTable = null)
        {
            this.Ping(msg: "$$$");

            COMTable = COMTable ?? _COMTable;

            var COMCells = COMTable.Range.Cells;//TODO unsafe

            var groupedRows = COMCells
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
            Log.Debug(DumpGrid(result, $"\n{nameof(GetMasterGrid)}-result"));

            this.Pong();

            return result;
        }

        Base1JaggedList<string> _textGrid = new Base1JaggedList<string>();
        internal Base1JaggedList<GridCell5> CrawlVertically(
            int rowOffset = default,
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {
            this.Ping(msg: "$$$");

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
                                gridRow.Insert(colIndex, new MergedGridCell5(up1Cell.GridRow + rowOffset, up1Cell.GridCol, up1Cell));
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



            Log.Debug(GridCrawler5.DumpGrid(textGrid, $"{nameof(CrawlVertically)}-{nameof(textGrid)}"));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, $"{nameof(CrawlVertically)}-{nameof(normalizedGrid)}"));

            this.Pong();

            return normalizedGrid;
        }





        internal Base1JaggedList<GridCell5> CrawlHoriz(
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {
            this.Ping(msg: "$$$");



            textGrid = textGrid ?? _textGrid;
            normalizedGrid = normalizedGrid ?? _grid;




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


            Log.Debug(GridCrawler5.DumpGrid(textGrid, $"{nameof(CrawlHoriz)}-{nameof(textGrid)}"));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, $"{nameof(CrawlHoriz)}-{nameof(normalizedGrid)}"));

            this.Pong();

            return normalizedGrid;
        }


        public static (Word.Table first, Word.Table second, int splitRow) SplitTableAtRow(Word.Table original)
        {
            if (original == null) throw new ArgumentNullException(nameof(original));

            var doc = original.Range.Document;

            // Middle row cell — safest
            var midCell = original.Range.Cells[original.Range.Cells.Count / 2];
            int splitRow = midCell.RowIndex;

            original.Split(splitRow);

            var tables = doc.Tables.Cast<Word.Table>()
                .OrderBy(t => t.Range.Start)
                .ToList();

            if (tables.Count < 2)
                throw new InvalidOperationException("Table split failed.");

            return (tables[0], tables[1], splitRow);
        }

        private int _analyzeTableRecursivelyDepth = 0;
        public Base1JaggedList<GridCell5> AnalyzeTableRecursively(Word.Table table = null, int rowOffset = 0)
        {
            this.Ping();
            table = table ?? _COMTable;
            if (table == null) throw new ArgumentNullException(nameof(table));
            Log.Debug($"Analyzing table at position {table.Range.Start}, rowOffset = {rowOffset}");

            bool hasMerges = false;
            try
            {
                _ = table.Rows[1];
                _ = table.Columns[1];
            }
            catch (COMException ex)
            {
                Log.Debug("Table access failed due to possible merges: " + ex.Message);
                hasMerges = true;
            }

            if (!hasMerges)
            {
                Log.Debug("No merges detected. Using fast GridCell5 crawl.");
                return this.Pong(() => BuildFastGridFromTable(table, rowOffset));
            }

            int cellCount = table.Range.Cells.Count;

            if (cellCount > 100 || _analyzeTableRecursivelyDepth++ > 10)
            {
                Log.Debug($"Table is large ({cellCount} cells). Splitting and recursing.");

                try
                {
                    var (first, second, splitRow) = SplitTableAtRow(table);

                    var firstGrid = AnalyzeTableRecursively(first, rowOffset);
                    var secondGrid = AnalyzeTableRecursively(second, rowOffset + splitRow - 1);

                    return this.Pong(() => MergeGrids(firstGrid, secondGrid));
                }
                catch (Exception ex)
                {
                    Log.Warning("Failed to split table. Falling back to slow crawl. " + ex.Message);
                    var result = BuildSlowGridFromTable(table, rowOffset);
                    --_analyzeTableRecursivelyDepth;
                    return this.Pong(() => result);
                }
            }
            else
            {
                Log.Debug("Table is small. Using slow full crawl.");
                var result = BuildSlowGridFromTable(table, rowOffset);
                --_analyzeTableRecursivelyDepth;
                return this.Pong(() => result);
            }
        }

        public Base1JaggedList<GridCell5> MergeGrids(Base1JaggedList<GridCell5> g1, Base1JaggedList<GridCell5> g2)
        {
            var merged = new Base1JaggedList<GridCell5>();
            foreach (var row in g1) merged.Add(row);
            foreach (var row in g2) merged.Add(row);

            return merged;
        }

        public Base1JaggedList<GridCell5> BuildFastGridFromTable(Word.Table table, int rowOffset = 0)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var grid = new Base1JaggedList<GridCell5>();

            for (int rowIdx_1 = 1; rowIdx_1 <= table.Rows.Count; rowIdx_1++)
            {
                var row = new Base1List<GridCell5>();
                int globalRow = rowOffset + rowIdx_1;

                for (int colIdx_1 = 1; colIdx_1 <= table.Columns.Count; colIdx_1++)
                {
                    row.Add(new GridCell5(globalRow, colIdx_1));
                }

                row.Add(new RowEndGridCell5());
                grid.Add(row);
            }

            //debug tagging
            table.Range.Cells[1].Range.Text = "FGC";

            return grid;
        }


        public Base1JaggedList<GridCell5> BuildSlowGridFromTable(Word.Table table, int rowOffset = 0)
        {
            this.Ping();
            if (table == null) throw new ArgumentNullException(nameof(table));

            var clonedTable = PrepareTable();
            var masterGrid = GetMasterGrid(clonedTable);
            var textGrid = ParseTableText(clonedTable);
            var normalGrid = NormalizeByWidth(masterGrid, rowOffset);
            var horizGrid = CrawlHoriz(textGrid, normalGrid);
            var vertGrid = CrawlVertically(rowOffset, textGrid, normalGrid);

            return this.Pong(() => vertGrid);
        }


    }
}
