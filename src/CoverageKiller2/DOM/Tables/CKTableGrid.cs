
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;

[assembly: InternalsVisibleTo("CoverageKiller2_Tests")]
namespace CoverageKiller2.DOM.Tables
{
    public class CKTableGrid
    {
        private static Dictionary<CKTable, CKTableGrid> _tableGrids = new Dictionary<CKTable, CKTableGrid>();
        //private CKTable _ckTable;
        //private Word.Table _comTable;
        internal Base1JaggedList<GridCell5> _grid;
        internal GridCrawler5 _crawler;

        // 🍽 Shared across all internal grid ops
        internal int RowCount => _grid.Count;
        internal int ColCount => _grid.LargestRowCount - 1;//-1 to account for end of row cell.

        public static CKTableGrid GetInstance(CKTable ckTable, Word.Table comTable)
        {
            var bypassDebug = true;
            if (!bypassDebug)
            {
                int tableNum = ckTable.Document.Tables.IndexOf(ckTable);
                Log.Debug($"Getting CKTableGrid Instance for table {tableNum}" +
                    $" of {ckTable.Document.Tables.Count} from document '{ckTable.Document.FileName}'");

                if (tableNum == -1 && Debugger.IsAttached) Debugger.Break();

            }

            _tableGrids.Keys.Where(r => r.IsOrphan).ToList()
                .ForEach(r => _tableGrids.Remove(r));

            if (_tableGrids.TryGetValue(ckTable, out CKTableGrid grid))
            {
                return grid;
            }

            Log.Debug($"Grid Instance not found for table; creating new.");
            grid = new CKTableGrid(ckTable);//, comTable);
            _tableGrids.Add(ckTable, grid);

            return grid;
        }

        /// <summary>
        /// Retrieves all master cells from the grid that fall within the rectangular area defined by the grid reference.
        /// </summary>
        /// <param name="gridRef">The cell reference bounds (inclusive, 1-based) to search within.</param>
        /// <returns>An enumerable of <see cref="GridCell5"/> instances that are master cells within the specified bounds.</returns>
        internal IEnumerable<GridCell5> GetMasterCells(CKGridCellRef gridRef)
        {
            this.Ping();
            Log.Debug($"MasterCell requested: [{gridRef.RowMin}:{gridRef.ColMin}] to [{gridRef.RowMin}:{gridRef.ColMax}]");

            var result = new List<GridCell5>();

            for (int row = gridRef.RowMin; row <= gridRef.RowMax; row++)
            {
                if (row < 1 || row > _grid.Count) continue;

                var currentRow = _grid[row];
                for (int col = gridRef.ColMin; col <= gridRef.ColMax; col++)
                {
                    if (col < 1 || col > currentRow.Count) continue;

                    var cell = currentRow[col];

                    Log.Verbose($"Inspecting cell at [{row},{col}]: type={cell.GetType().Name}, master={cell.IsMasterCell}");
                    if (cell.IsMasterCell)
                    {
                        result.Add(cell);
                    }
                }
            }

            Log.Debug($"Found {result.Count} master cells.");
            if (!result.Any())
            {
                if (Debugger.IsAttached) Debugger.Break();
                throw new Exception("No master cells found.");
            }
            this.Pong();
            return result;
        }
        public bool HasMerge
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
            shadowWorkspace.ShowDebuggerWindow();

            //put original table
            shadowWorkspace.CloneFrom(sourceTable); //make sure we aren't recursing tables here.

            var x = shadowWorkspace.Document.Content.CollapseToEnd();
            x.COMRange.InsertAfter("\r\r\r");
            //put the one to format
            var clonedTable = shadowWorkspace.CloneFrom(sourceTable);
            //var grid = GetMasterGrid(clonedTable);
            //Log.Debug(GridCrawler5.DumpGrid(grid));

            //pulling once

            return this.Pong(() => clonedTable.COMTable);
        }
    }
}
