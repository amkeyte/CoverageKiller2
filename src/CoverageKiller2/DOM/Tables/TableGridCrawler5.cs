using CoverageKiller2.Logging;
using Serilog;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{

    /// <summary>
    /// Builds a reusable layout grid from a Word table using both visual and textual merge inference.
    /// </summary>
    /// <remarks>
    /// THIS CLASS IS NOT TO INCLUDE ANY CK.DOM references. Version: CK2.00.01.0034
    /// </remarks>
    public class GridCrawler5
    {
        private readonly Word.Table _COMTable;
        private Base1JaggedList<GridCell5> _grid = new Base1JaggedList<GridCell5>();
        private Base1JaggedList<GridCell5> _masterCells = new Base1JaggedList<GridCell5>();
        //private ShadowWorkspace _shadowWorkspace;
        internal GridCrawler5(Word.Table table)
        {
            _COMTable = table ?? throw new ArgumentNullException(nameof(table));
            //_shadowWorkspace = workspace ?? throw new ArgumentNullException(nameof(workspace));
            _grid = AnalyzeTableRecursively(table);

            Log.Debug(DumpGrid(_grid, $"\n\n***  Final table for Table:" +
                $"[{LH.GetTableTitle(_COMTable, "***Table")}] [{new RangeSnapshot(table.Range).FastHash}]***"));

            if (_grid.SelectMany(r => r).Any(c => c.IsGhostCell)) throw new CKDebugException("Crawl failed Ghost cells not resolved.");


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
        ///// <summary>
        ///// Returns a labeled text representation of the grid.
        ///// </summary>
        //public static string DumpGrid(Base1JaggedList<string> grid, string message = null)
        //{
        //    var sb = new StringBuilder();
        //    sb = sb.AppendLine(message + "\n");
        //    sb.AppendLine("**********************");

        //    foreach (var row in grid)
        //    {
        //        var line = row.Select(cell => string.IsNullOrEmpty(cell) ? "[NULL]" : cell);
        //        sb.AppendLine(string.Join(" | ", line));
        //    }
        //    sb.AppendLine("**********************");

        //    return sb.ToString();
        //}
        /// <summary>
        /// Returns a labeled text representation of a Base1JaggedList of any type.
        /// </summary>
        /// <typeparam name="T">The element type of the grid.</typeparam>
        /// <param name="grid">The grid to dump.</param>
        /// <param name="message">Optional label message.</param>
        /// <returns>A formatted string representation of the grid.</returns>
        public static string DumpGrid<T>(Base1JaggedList<T> grid, string message = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine(message + "\n");
            sb.AppendLine("**********************");

            foreach (var row in grid)
            {
                var line = row.Select(cell => cell?.ToString() ?? "[NULL]");
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

        //    //LH.Checkpoint($"Cloning table {_table.Document.Tables.IndexOf(_table)} from document {_table.Document.FileName}");

        //    //fast crawl

        //    //slow crawl
        //    var clonedTable = PrepareTable();
        //    var masterGrid = GetMasterGrid(clonedTable);
        //    var textGrid = ParseTableText(clonedTable);
        //    var normalGrid = NormalizeByWidth(masterGrid);
        //    var horizGrid = CrawlHoriz(textGrid, normalGrid);
        //    var vertGrid = CrawlVertically(0, textGrid, normalGrid);

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
        /// Parses the table's raw emitted text into a jagged list of cell strings,
        /// accounting for Word’s \r\a patterns and logical row structure.
        /// </summary>
        /// <param name="COMTable">Optional override of the source table. Defaults to internal _COMTable.</param>
        /// <returns>A jagged list of strings representing logical cell values.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0037
        /// </remarks>
        internal Base1JaggedList<string> ParseTableText(Word.Table COMTable = null)
        {

            // Use default if not provided
            COMTable = COMTable ?? _COMTable;
            var COMRange = COMTable.Range;

            // Determine visual width to normalize grid width later
            var colCount = GetMasterGrid(COMTable).LargestRowCount;

            // Split the raw table text by Word's row delimiters (\r\a\r\a)
            var rawText = COMRange.Text;
            var rowTexts = SplitWordTableTextIntoRows(rawText);
            Log.Verbose(DumpList(rowTexts, "Row Texts"));

            var result = new Base1JaggedList<string>();

            // Each row is further split into cell text chunks using \r\a
            foreach (var row in rowTexts)
            {
                var matches = Regex.Matches(row, @"(.*?\r\a)")
                                   .Cast<Match>()
                                   .Select(m => FlattenTableText(m.Value))
                                   .ToList();

                result.Add(new Base1List<string>(matches));
            }

            // Normalize column widths by forward-shuffling from below rows
            for (int i = 1; i < result.Count; i++)
            {
                var row = result[i];
                var nextRowIndex = i + 1;

                while (row.Count < colCount + 1 && result.Count >= nextRowIndex)
                {
                    row.Add(result[nextRowIndex][1]);
                    result[nextRowIndex].RemoveAt(1);
                    if (result[nextRowIndex].Count == 0)
                        result.RemoveAt(nextRowIndex);
                }
            }

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

            var result = new Base1JaggedList<GridCell5>();
            var widestRow = masterGrid
             .Zip(_textGrid, (gridRow, textRow) => new { gridRow, textRow })
             .Where(x => x.textRow.Count == x.gridRow.Count)
             .OrderByDescending(x => x.gridRow.Sum(c => c.Width))
             .Select(x => x.gridRow)
             .FirstOrDefault();

            // Fallback if no good match found
            if (widestRow == null) widestRow = masterGrid[2];

            float totalRowWidth = widestRow.Sum(c => c.Width);
            int colCount = widestRow.Count;
            float normalWidth = totalRowWidth / colCount;

            //Log.Debug($"[NormalizeByWidth] Using matched row width = {totalRowWidth}, columns = {colCount}");
            //Log.Debug($"[NormalizeByWidth] Calculated normalWidth = {normalWidth}");


            // Step 4: Pad each row by inserting ZombieCells after wide master cells
            var newGrid = new Base1JaggedList<GridCell5>();

            foreach (var row in masterGrid)
            {
                var newRow = new Base1List<GridCell5>();
                //insert cells where there are wide spaces
                var _debugNewGridCellAddedCount = 0;
                foreach (var cell in row)
                {

                    //Log.Debug($"Row {cell.RowIndex}, Col {cell.ColumnIndex} width = {cell.Width}");
                    var newCell = new GridCell5(cell.RowIndex + rowOffset, cell.ColumnIndex);
                    _debugNewGridCellAddedCount++;
                    newRow.Add(newCell);

                    var _debugNewMergedCellAddedCount = 0;
                    int span = Math.Max(1, (int)Math.Round(cell.Width / normalWidth));

                    for (int i = 1; i < span; i++)//for one cell, span is 1. 
                    {
                        newRow.Add(new MergedGridCell5(cell.RowIndex + rowOffset, cell.ColumnIndex + i, newCell));
                        _debugNewMergedCellAddedCount++;
                    }

                    //Log.Debug($"\n\nRow {cell.RowIndex}, Col {cell.ColumnIndex} width = {cell.Width}\n\t" +
                    //    $"Added new cells to cover span of {span}: GridCell ({_debugNewGridCellAddedCount}) MegedCell {_debugNewMergedCellAddedCount}");
                }
                //insert cells to fill out row
                for (int i = newRow.Count; i < colCount; i++)
                {
                    newRow.Add(new GhostGridCell5());
                }
                newRow.Add(new RowEndGridCell5());
                newGrid.Add(newRow);
            }
            //Log.Debug(DumpGrid(newGrid, $"\n{nameof(NormalizeByWidth)}-{nameof(newGrid)}"));

            _grid = newGrid;

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

            COMTable = COMTable ?? _COMTable;
            var COMRange = _COMTable.Range;

            ///Bug 20250426-0026 fixed by stabling the table column spacing
            ///using the normailization techniques below.

            //*******************************
            // here we normalize all the table's settings so that
            // the table will produce the correct numbers for measurement.

            //normalize the text
            COMRange.Font.Name = "Consolas";
            COMRange.Font.Size = 10;
            // Clear out extra spacing
            COMTable.TopPadding = 0f;
            COMTable.BottomPadding = 0f;
            COMTable.LeftPadding = 0f;
            COMTable.RightPadding = 0f;
            // Remove any cell spacing (sometimes tables have inside spacing between cells)
            COMTable.Spacing = 0f;
            // Remove table indentation from the left margin
            COMTable.Rows.LeftIndent = 0f;
            // Optional: Normalize individual cell padding too
            foreach (Word.Cell cell in COMTable.Range.Cells)
            {
                cell.TopPadding = 0f;
                cell.BottomPadding = 0f;
                cell.LeftPadding = 0f;
                cell.RightPadding = 0f;
            }

            //*******************************
            // Gere we adjust the table's width and column behavior
            // so that the columns auto-adjust into equal segments.
            // this is important because the layout of horizontal
            // merges will now become sized as multiples of the 
            // regular grid size.

            // Enable auto-fit so columns stretch evenly to fill the table width
            COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            // set the table width to make.
            COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            COMTable.PreferredWidth = 100f;


            //get the longest row (assuming to have all the columns)
            var longestRow = GetMasterGrid(COMTable)
                .OrderByDescending(r => r.Count)
                .FirstOrDefault();

            // ~78 X's will fit across the page. by stuffing each cell full 
            // to word wrap, it should autosize each column an equal width.
            // this will help avoid a trap where autosze can be inconsistent causing phantom hits
            // on the width/span algorythm.
            string cellStretcher = new string('X', 100 / longestRow.Count);
            COMTable.Range.Text = string.Empty;

            foreach (Word.Cell cell in COMRange.Cells)
            {
                cell.Range.Text = cellStretcher;
            }

            CKDocument.EnsureLayoutReady(COMTable.Range.Document);

            // after the adjustments are made, lock in the fitting to prevent
            // movement in the future.
            COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);


            return COMTable;
        }



        /// <summary>
        /// Retrieves a grid of master GridCell5s from the CKTable.
        /// </summary>
        public Base1JaggedList<Word.Cell> GetMasterGrid(Word.Table COMTable = null)
        {

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
            //Log.Verbose(DumpGrid(result, $"\n{nameof(GetMasterGrid)}-result"));

            _COMCellGrid = result;

            return result;
        }


        /// <summary>
        /// Infers vertical cell merging by analyzing vertical alignment between the normalized grid and text grid.
        /// Updates the layout grid in-place to reflect vertical merges (row spans).
        /// </summary>
        /// <param name="rowOffset">Starting row offset in global coordinates.</param>
        /// <param name="textGrid">Text representation of cells, defaults to _textGrid if null.</param>
        /// <param name="normalizedGrid">Normalized visual grid from CrawlHoriz, defaults to _grid if null.</param>
        /// <returns>The vertically merged visual grid.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0039
        /// </remarks>
        internal Base1JaggedList<GridCell5> CrawlVertically(
            int rowOffset = 0,
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {

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
                        // If cell is visually empty but not marked merged, something went wrong
                        if (textCell != "/r/a") continue;
                    }
                    else if (gridCell.IsMergedCell)
                    {
                        // Confirm text grid alignment
                        if (textCell == "/r/a") continue;
                    }
                    else if (gridCell.IsGhostCell)
                    {
                        // This cell visually aligns with above master; replace with merged cell(s)
                        if (textCell != "/r/a") throw new Exception("Ghost cell expected to be empty");

                        var up1Index = rowIndex - 1;
                        if (up1Index < 1)
                        {
                            Log.Warning("Ghost cell at top row cannot be resolved — skipping.");
                            continue;
                        }

                        var up1Row = normalizedGrid[up1Index];
                        var up1Cell = up1Row[colIndex];

                        // Replace ghost cell with merged cell block
                        gridRow.RemoveAt(colIndex);

                        if (up1Cell.IsMasterCell)
                        {
                            for (int i = 0; i < up1Cell.ColSpan; i++)
                            {
                                gridRow.Insert(colIndex, new MergedGridCell5(
                                    up1Cell.GridRow + rowOffset,
                                    up1Cell.GridCol,
                                    up1Cell));
                            }
                        }
                        else if (up1Cell.IsMergedCell)
                        {
                            for (int i = 0; i < up1Cell.MasterCell.ColSpan; i++)
                            {
                                gridRow.Insert(colIndex,
                                    new MergedGridCell5(
                                        up1Cell.MasterCell.GridRow,
                                        up1Cell.MasterCell.GridCol,
                                        up1Cell.MasterCell));
                            }
                        }
                        else
                        {
                            throw new Exception("Unexpected upstream cell type above ghost");
                        }
                    }
                    else if (gridCell.IsRowEndMarker)
                    {
                        // End marker; just continue
                        continue;
                    }
                }
            }

            _grid = normalizedGrid;

            //Log.Verbose(DumpGrid(textGrid, $"{nameof(CrawlVertically)}-{nameof(textGrid)}"));
            //Log.Verbose(DumpGrid(normalizedGrid, $"{nameof(CrawlVertically)}-{nameof(normalizedGrid)}"));

            return normalizedGrid;
        }

        private Base1JaggedList<string> _textGrid;
        internal Base1JaggedList<GridCell5> CrawlHoriz(
            Base1JaggedList<string> textGrid = null,
            Base1JaggedList<GridCell5> normalizedGrid = null)
        {




            // Use previously cached grids if not supplied
            textGrid = textGrid ?? _textGrid;
            normalizedGrid = normalizedGrid ?? _grid;



            // Determine the widest row to iterate columns safely
            var normalizedRowCount = normalizedGrid.LargestRowCount;

            // Loop through each row
            for (var rowIndex = 1; rowIndex <= normalizedGrid.Count; rowIndex++)
            {
                var gridRow = normalizedGrid[rowIndex];   // visual grid row of GridCell5
                var textRow = textGrid[rowIndex];         // raw text values for same row

                // Loop across the widest row column count
                for (var cellIndex = 1; cellIndex <= normalizedRowCount; cellIndex++)
                {
                    var gridCell = normalizedGrid.SafeGet(rowIndex, cellIndex);     // GridCell5 at this row,col
                    var textCell = textGrid.SafeGet(rowIndex, cellIndex);    // Corresponding text from ParseTableText()
                    if (gridCell == null || textCell == null)
                    {
                        Log.Error($"CrawlHoriz mismatch at ({rowIndex},{cellIndex}) - " +
                                 $"normalized: {(gridCell == null ? "null" : "OK")}, " +
                                 $"text: {(textCell == null ? "null" : "OK")}");

                        throw new CKDebugException($"Abnormal cell in CrawlHoriz at ({rowIndex},{cellIndex})");
                    }

                    // If it's a master cell and the text shows a merged marker (/r/a)
                    if (gridCell.IsMasterCell)
                    {
                        // If the text value isn't a merge indicator, skip
                        if (textCell != "/r/a") continue;

                        // Otherwise, insert a ghost cell to realign the layout grid
                        gridRow.Insert(cellIndex, new GhostGridCell5());
                    }

                    // If it's a merged cell, we expect a "/r/a" in the text
                    else if (gridCell.IsMergedCell)
                    {
                        // If the merge marker is already there, no fix needed
                        if (textCell == "/r/a") continue;

                        // Otherwise, insert the merge marker into the text row
                        textRow.Insert(cellIndex, "/r/a");
                    }

                    // Handle row end markers (used to force row alignment)
                    else if (gridCell.IsRowEndMarker)
                    {
                        // If the text doesn't end with "/r/a", patch it in
                        if (textCell != "/r/a") textRow.Insert(cellIndex, "/r/a");

                        // Now push all cells past the row end to the next row
                        for (var i = textRow.Count; i > cellIndex; i--)
                        {
                            // Add new row if it doesn't exist yet
                            if (rowIndex + 1 > textGrid.Count)
                                textGrid.Add(new Base1List<string>());

                            // Move last cell in this row to the front of the next row
                            textGrid[rowIndex + 1].Insert(1, textRow.Last());

                            // Remove the moved cell from current row
                            textRow.RemoveAt(textRow.Count);
                        }
                    }

                    // Ghost cells are placeholders introduced earlier
                    else if (gridCell.IsGhostCell)
                    {
                        // If it's beyond the expected width, drop it and rewind index
                        if (cellIndex >= normalizedRowCount)
                        {
                            gridRow.RemoveAt(cellIndex);
                            cellIndex--;
                            continue;
                        }

                        ///added to debug 20250425-0013***********
                        // If text row is too short, pad it out so we can safely access this index
                        while (textRow.Count < cellIndex)
                        {
                            textRow.Add("/r/a"); // or null?
                        }
                        //***************************************

                        // If the text aligns, ghost is valid
                        if (textCell == "/r/a") continue;

                        // Otherwise, align text row by inserting the ghost cell marker
                        textRow.Insert(cellIndex, "/r/a");
                    }
                }
            }


            //validate
            if (textGrid.Flatten().Count() != normalizedGrid.Flatten().Count())
                throw new ArgumentException("Input grids are out of alignment.");


            // Save mutated structures back into local cache
            _textGrid = textGrid;
            _grid = normalizedGrid;

            // Output final structures for debugging
            //Log.Verbose(DumpGrid(textGrid, $"{nameof(CrawlHoriz)}-{nameof(textGrid)}"));
            //Log.Verbose(DumpGrid(normalizedGrid, $"{nameof(CrawlHoriz)}-{nameof(normalizedGrid)}"));


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
        private Base1JaggedList<Word.Cell> _COMCellGrid;

        public Base1JaggedList<GridCell5> AnalyzeTableRecursively(Word.Table table = null, int rowOffset = 0)
        {
            table = table ?? _COMTable;
            if (table == null) throw new ArgumentNullException(nameof(table));
            //Log.Debug($"Analyzing table at position {table.Range.Start}, rowOffset = {rowOffset}");

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
                return BuildFastGridFromTable(table, rowOffset);
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

                    return MergeGrids(firstGrid, secondGrid);
                }
                catch (Exception ex)
                {
                    Log.Warning("Failed to split table. Falling back to slow crawl. " + ex.Message);
                    var result = BuildSlowGridFromTable(table, rowOffset);
                    --_analyzeTableRecursivelyDepth;
                    return result;
                }
            }
            else
            {
                Log.Debug("Table is small. Using slow full crawl.");
                var result = BuildSlowGridFromTable(table, rowOffset);
                --_analyzeTableRecursivelyDepth;
                return result;
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

            try
            {

                if (table == null) throw new ArgumentNullException(nameof(table));

                var clonedTable = PrepareTable();
                var masterGrid = GetMasterGrid(clonedTable);
                var textGrid = ParseTableText(clonedTable);
                var normalGrid = NormalizeByWidth(masterGrid, rowOffset);
                var horizGrid = CrawlHoriz(textGrid, normalGrid);
                var vertGrid = CrawlVertically(rowOffset, textGrid, normalGrid);

                return vertGrid;
            }
            catch (Exception ex)
            {
                //diagnostics

                Log.Error(ex, "Exception in Horizonatal crawl.");

                //concoct the width grid data***

                var widthGrid = new Base1JaggedList<float>();

                foreach (var row in _COMCellGrid)
                {
                    var widthRow = new Base1List<float>();
                    foreach (var cell in row)
                    {
                        widthRow.Add(cell.Width);
                    }
                    widthGrid.Add(widthRow);
                }
                var totalWidth = widthGrid.LargestRow.Sum();
                var normalWidth = totalWidth / widthGrid.LargestRowCount;
                // *** done


                //cast the log dump
                Log.Debug("Diagnostic info:");
                var sb = new StringBuilder();
                sb.AppendLine(DumpGrid(_COMCellGrid, $"-{nameof(_COMCellGrid)}"));
                sb.AppendLine(DumpGrid(widthGrid, $"-{nameof(widthGrid)}"));
                sb.AppendLine($"Total Table width: {totalWidth}, Normal cell size: {normalWidth}");
                sb.AppendLine(DumpGrid(_textGrid, $"-{nameof(_textGrid)}"));
                sb.AppendLine(DumpGrid(_grid, $"-{nameof(_grid)}"));

                //summon the CKDocument
                var wdDoc = _COMTable.Range.Document;
                if (CKOffice_Word.Instance.TryGetAppFor(wdDoc, out CKApplication ckApp)
                    && ckApp.TryGetDocumentFor(wdDoc, out CKDocument ckDoc))
                {
                    ckDoc.Visible = true;
                    ckDoc.Activate();
                    ckDoc.ActiveWindow.Activate();
                    ckDoc.KeepAlive = true;
                    sb.AppendLine($"Document {ckDoc.FileName} is available for review.");
                }
                else
                {
                    sb.AppendLine($"Unable to make {Path.GetFileName(wdDoc.FullName)} visible.");
                }

                LH.Debug(sb.ToString());

                throw ex;
            }
        }


    }

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


}
