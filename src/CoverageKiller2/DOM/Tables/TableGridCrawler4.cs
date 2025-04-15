using CoverageKiller2.Logging;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Successor utility class to TableGridCrawler3 for visualizing and coloring Word table grids.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0014
    /// </remarks>
    public static class TableGridCrawler4
    {
        public static Tracer Tracer = new Tracer(typeof(TableGridCrawler4));

        /// <summary>
        /// Dumps a string representation of the grid for debugging or visualization.
        /// </summary>
        /// <param name="grid">The jagged grid structure of GridCell2 instances.</param>
        /// <returns>A formatted string showing grid positions and cell types.</returns>
        public static string DumpGrid(Base1JaggedList<GridCell2> grid)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("\n");

            foreach (var row in grid)
            {
                var line = row.Select(c =>
                {
                    var label = c.IsDummy ? "Z" : (c.IsMerged ? "M" : "O");
                    return $"{label}[{c.MasterCell.GridRow},{c.MasterCell.GridCol}]";
                });

                sb.AppendLine(string.Join(" | ", line));
            }

            return sb.ToString();
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
        /// Constructs a jagged list of master GridCell2 instances from a CKTable.
        /// Only includes top-left anchors of merged regions.
        /// </summary>
        /// <param name="table">The CKTable to extract from.</param>
        /// <returns>A jagged grid of master GridCell2s.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0017
        /// </remarks>
        public static Base1JaggedList<GridCell2> GetMasterGrid(CKTable table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var groupedRows = table.COMTable.Range.Cells
                .Cast<Word.Cell>()
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(c => c.ColumnIndex))
                .ToList();

            var result = new Base1JaggedList<GridCell2>();

            foreach (var row in groupedRows)
            {
                var list = new Base1List<GridCell2>();
                foreach (var cell in row)
                {

                    list.Add(new GridCell2(cell, cell.RowIndex, cell.ColumnIndex, isMasterCell: true));
                }
                result.Add(list);
            }

            return result;
        }
        /// <summary>
        /// Creates a new table at the end of the shadow document using the provided visual grid.
        /// Each cell is labeled with its [row,col] coordinates.
        /// </summary>
        /// <param name="location">The shadow workspace to insert into.</param>
        /// <param name="grid">The visual grid to translate into a table.</param>
        /// <returns>The inserted CKTable representing the grid.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0026
        /// </remarks>
        public static CKTable InsertGridAsTableAtEnd(IDOMObject location, Base1JaggedList<GridCell2> grid)
        {
            if (location == null) throw new ArgumentNullException(nameof(location));
            if (grid == null || grid.Count == 0) throw new ArgumentException("Grid is null or empty.", nameof(grid));

            var doc = location.Document;
            var app = location.Application;

            // Step 1: create an insertion range at the end
            var insertRange = doc.Range().CollapseToEnd();

            int rowCount = grid.Count;
            int colCount = grid.Max(row => row.Count);

            // Step 2: insert new Word table
            var newTable = doc.Tables.Add(
                 insertRange,
                 rowCount,
                 colCount);

            // Step 3: apply font and width formatting
            newTable.COMRange.Font.Name = "Consolas";
            newTable.COMRange.Font.Size = 10;

            // Enable auto-fit so columns stretch evenly to fill the table width
            newTable.COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

            newTable.COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            newTable.COMTable.PreferredWidth = 100f;

            // Step 4: fill in the table using GridCell2's actual GridRow/GridCol values
            for (int r = 1; r <= rowCount; r++)
            {
                var sourceRow = grid[r];
                for (int c = 1; c <= colCount; c++)
                {
                    if (c <= sourceRow.Count)
                    {
                        var sourceCell = sourceRow[c];

                        var destCell = newTable.COMTable.Cell(r, c);
                        destCell.Range.Text = $"[{sourceCell.MasterCell.GridRow},{sourceCell.MasterCell.GridCol}]";
                    }
                }
            }

            // Step 5: wrap in CKTable and return
            return newTable;
        }

        /// <summary>
        /// Builds and normalizes a GridCell2 layout from a CKTable using actual Word cell widths.
        /// Pads merged cells visually by inserting ZombieCell2 instances based on width ratios.
        /// </summary>
        /// <param name="table">The CKTable to process.</param>
        /// <returns>A normalized, padded GridCell2 jagged grid.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0025
        /// </remarks>
        public static Base1JaggedList<GridCell2> NormalizeGridByMeasuredWidth(CKTable table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));



            // Step 1: Build master-only grid
            var masterGrid = GetMasterGrid(table);
            Tracer.Log($"Dump Grid {nameof(masterGrid)}: \n{TableGridCrawler4.DumpGrid(masterGrid)}\n");


            if (masterGrid.Count == 0)
                throw new InvalidOperationException("Table has no master grid rows.");

            // Step 2: Find the row with the most master cells
            var widestRow = masterGrid.OrderByDescending(r => r.Count).First();
            float totalRowWidth = widestRow.Sum(c => c.COMCell.Width);
            int colCount = widestRow.Count;
            // Step 3: Average width per column (baseline)
            float normalWidth = totalRowWidth / colCount;

            // Step 4: Pad each row by inserting ZombieCells after wide master cells
            var newGrid = new Base1JaggedList<GridCell2>();

            foreach (var row in masterGrid)
            {
                var newRow = new Base1List<GridCell2>();

                foreach (var cell in row)
                {
                    newRow.Add(cell);

                    int span = Math.Max(1, (int)Math.Round(cell.COMCell.Width / normalWidth));


                    for (int i = 1; i <= span - 1; i++)
                    {
                        newRow.Add(new ZombieCell2(cell, cell.GridRow, cell.GridCol + i));
                    }
                }

                newGrid.Add(newRow);
            }
            Tracer.Log(DumpGrid(newGrid));
            return newGrid;
        }


        /// <summary>
        /// Clones the given Word table into a shadow workspace and formats it for grid-based layout visualization.
        /// </summary>
        /// <param name="source">The table to clone and prepare.</param>
        /// <param name="colWidth">Optional column width in points (default: 20f).</param>
        /// <returns>A ShadowWorkspace containing the formatted cloned table and master cell grid.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0015
        /// </remarks>
        public static CKTable CloneAndPrepareTableLayout(
            CKTable source,
            float colWidth = 20f)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            var app = CKOffice_Word.Instance.Applications.FirstOrDefault();
            var workspace = app.GetShadowWorkspace();
            var clonedTable = workspace.CloneFrom(source);
            var grid = GetMasterGrid(clonedTable);
            LabelCellsWithCoordinates(grid);



            try
            {
                //var table = grid.Get2D(1, 1).COMCell.Range.Tables[1];
                clonedTable.COMRange.Font.Name = "Consolas";
                clonedTable.COMRange.Font.Size = 10;

                // Enable auto-fit so columns stretch evenly to fill the table width
                clonedTable.COMTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                clonedTable.COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                clonedTable.COMTable.PreferredWidth = 100f;


                workspace.ShowDebuggerWindow();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return clonedTable;
        }
        /// <summary>
        /// Overwrites the text in each master cell with its [row,col] coordinates using LINQ.
        /// </summary>
        /// <param name="grid">The grid to annotate.</param>
        /// <remarks>
        /// Version: CK2.00.01.0019
        /// </remarks>
        public static void LabelCellsWithCoordinates(Base1JaggedList<GridCell2> grid)
        {
            grid
                .SelectMany(row => row)
                .ToList()
                .ForEach(cell =>
                {
                    try
                    {
                        cell.COMCell.Range.Text = $"[{cell.MasterCell.GridRow},{cell.MasterCell.GridCol}]";
                    }
                    catch
                    {
                        // Ignore safely
                    }
                });
        }


    }
}
