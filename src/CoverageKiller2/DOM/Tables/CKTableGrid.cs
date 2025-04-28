
using CoverageKiller2.Logging;
using Serilog;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;

[assembly: InternalsVisibleTo("CoverageKiller2_Tests")]
namespace CoverageKiller2.DOM.Tables
{
    public class CKTableGrid
    {
        private static Dictionary<string, CKTableGrid> _tableGrids = new Dictionary<string, CKTableGrid>();
        //private CKTable _ckTable;
        //private Word.Table _comTable;
        internal Base1JaggedList<GridCell5> _grid;
        internal GridCrawler5 _crawler;

        // 🍽 Shared across all internal grid ops
        internal int RowCount => _grid.Count;
        internal int ColCount => _grid.LargestRowCount - 1;//-1 to account for end of row cell.

        public static CKTableGrid GetInstance(CKTable ckTable, Word.Table comTable, [MemberCallerName] string callerName = null)
        {
            LH.Ping<CKTableGrid>();
            var tableId = $"{ckTable.Document.FileName}[{ckTable.Snapshot.FastHash}]";
            Log.Debug($"Getting CKTableGrid Instance for [{LH.GetTableTitle(ckTable, "***Table")}] {tableId} :: " +
                $"\n\t\t\t\t" +
                $"Requested by{ckTable.Parent.GetType()}::{ckTable.GetType()}::{callerName}.");

            if (_tableGrids.TryGetValue(tableId, out CKTableGrid grid))
            {
                LH.Pong<CKTableGrid>();

                return grid;
            }

            //Log.Debug($"Grid Instance not found for table; creating new.");
            grid = new CKTableGrid(ckTable);//, comTable);
            _tableGrids.Add(tableId, grid);

            //Log.Debug(new Base1List<string>(_tableGrids.Keys.ToList()).Dump("Available instances:"));
            LH.Pong<CKTableGrid>();

            return grid;
        }
        internal IEnumerable<GridCell5> GetMergedCells(CKGridCellRef gridRef)
        {
            Log.Debug($"MergedCells requested for: [{gridRef.RowMin}:{gridRef.ColMin}] to [{gridRef.RowMax}:{gridRef.ColMax}]");

            var result = new List<GridCell5>();

            for (int row_1 = gridRef.RowMin; row_1 <= gridRef.RowMax; row_1++)
            {
                if (row_1 < 1 || row_1 > RowCount) continue;

                var currentRow = _grid[row_1];
                for (int col = gridRef.ColMin; col <= gridRef.ColMax; col++)
                {
                    if (col < 1 || col > currentRow.Count) continue;

                    var cell = currentRow[col];

                    Log.Verbose($"Inspecting cell at [{row_1},{col}]: type={cell.GetType().Name}");

                    if (cell.IsMergedCell)
                    {
                        result.Add(cell);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Retrieves all master cells from the grid that fall within the rectangular area defined by the grid reference.
        /// </summary>
        /// <param name="gridRef">The cell reference bounds (inclusive, 1-based) to search within.</param>
        /// <returns>An enumerable of <see cref="GridCell5"/> instances that are master cells within the specified bounds.</returns>
        internal IEnumerable<GridCell5> GetMasterCells(CKGridCellRef gridRef)
        {
            Log.Debug($"MasterCells requested for: [{gridRef.RowMin}:{gridRef.ColMin}] to [{gridRef.RowMin}:{gridRef.ColMax}]");

            var result = new List<GridCell5>();

            for (int row_1 = gridRef.RowMin; row_1 <= gridRef.RowMax; row_1++)
            {
                if (row_1 < 1 || row_1 > RowCount) continue;

                var currentRow = _grid[row_1];
                for (int col = gridRef.ColMin; col <= gridRef.ColMax; col++)
                {
                    if (col < 1 || col > currentRow.Count) continue;

                    var cell = currentRow[col];

                    Log.Verbose($"Inspecting cell at [{row_1},{col}]: type={cell.GetType().Name}");
                    if (cell.IsMasterCell)
                    {
                        result.Add(cell);
                    }
                    else if (cell.IsMergedCell)
                    {
                        result.Add(cell.MasterCell);
                    }
                }
            }

            Log.Debug($"Found {result.Count} master cells.");
            if (!result.Any())
            {
                throw new CKDebugException("No master cells found.");
            }
            this.Pong();
            return result;
        }
        public bool HasMerge//TODO cache this someday
        {
            get
            {
                var result = _grid
                    .Where(row => row != null)
                    .SelectMany(row => row.Where(cell => cell != null))
                    .Any(cell => cell.IsMergedCell);

                return result;
            }
        }

        private CKTableGrid(CKTable parent)//, Word.Table table)
        {
            this.Ping(msg: parent.Snapshot.FastHash.ToString());
            //_ckTable = parent;
            //_comTable = table;
            var clonedTable = CloneToShadow(parent, parent.Application.GetShadowWorkspace());
            _crawler = new GridCrawler5(clonedTable);
            _grid = _crawler.Grid;
            this.Pong();
        }

        private Word.Table CloneToShadow(CKTable sourceTable, ShadowWorkspace shadowWorkspace)
        {
            this.Ping();
            //for debugging uncomment.
            //shadowWorkspace.ShowDebuggerWindow();
            //shadowWorkspace.Document.KeepAlive = true;
            //shadowWorkspace.Document.ActiveWindow.Activate();

            //put original table
            shadowWorkspace.CloneFrom(sourceTable); //make sure we aren't recursing tables here.

            var x = shadowWorkspace.Document.Content.CollapseToEnd();
            x.COMRange.InsertAfter("\r\r\r");
            //put the one to format
            var clonedTable = shadowWorkspace.CloneFrom(sourceTable, x.CollapseToEnd());
            //var grid = GetMasterGrid(clonedTable);
            //Log.Debug(GridCrawler5.DumpGrid(grid));

            //pulling once

            return this.Pong(() => clonedTable.COMTable);
        }
    }
}
