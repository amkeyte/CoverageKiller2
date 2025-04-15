using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    public class TableGridCrawler
    {
        private readonly Word.Table _table;

        public TableGridCrawler(Word.Table table)
        {
            _table = table ?? throw new ArgumentNullException(nameof(table));
        }
        public Word.Cell GetBottomRightCell()
        {
            Word.Cell bottomRight = null;
            int maxRow = -1;
            int maxCol = -1;

            foreach (Word.Cell cell in _table.Range.Cells)
            {
                try
                {
                    int row = cell.RowIndex;
                    int col = cell.ColumnIndex;

                    if (row > maxRow || (row == maxRow && col > maxCol))
                    {
                        maxRow = row;
                        maxCol = col;
                        bottomRight = cell;
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Skip phantom cells. Word loves its ghosts.
                    continue;
                }
            }

            return bottomRight;
        }
        public Base1JaggedList<GridCell> CrawlRows()
        {
            return CrawlDirection(
                crawlToVert: 0,
                crawlToHoriz: 1,
                makeGridCell: (row, col, cell) =>
                    new GridCell(cell, row, col, true),
                true
            );
        }
        public Base1JaggedList<GridCell> CrawlRowsReverse()
        {
            return CrawlDirection(
                crawlToVert: 0,
                crawlToHoriz: -1,
                makeGridCell: (row, col, cell) =>
                    new GridCell(cell, row, col, true),
                true
            );
        }

        public Base1JaggedList<GridCell> CrawlColumns()
        {
            return CrawlDirection(
                crawlToVert: 1,
                crawlToHoriz: 0,
                //a bit unclear, but okay I get it for now.
                makeGridCell: (row, col, cell) =>
                    new GridCell(cell, row, col, true)
            );
        }


        private Base1JaggedList<GridCell> CrawlDirection(
            int crawlToVert,
            int crawlToHoriz,
            Func<int, int, Word.Cell, GridCell> makeGridCell,
            bool reverse = false)
        {
            var visitedFastHashes = new HashSet<ulong>();//track what cells have been seen
            var outputLists = new Base1JaggedList<GridCell>();
            var wordAppCells = new Base1List<Word.Cell>(_table.Range.Cells.ToList());
            if (reverse) { wordAppCells.Reverse(); }

            foreach (var wordAppCell in wordAppCells)
            {
                //if (visitedFastHashes.Contains(snap.FastHash))
                //{
                //    //does this actually ever happen?
                //    if (Debugger.IsAttached) Debugger.Break();
                //    continue;
                //}

                //int wordTableRow = wordAppCell.RowIndex;
                //int wordTableCol = wordAppCell.ColumnIndex;

                //var thisGridCell = makeGridCell(wordTableRow, wordTableCol, wordAppCell);
                var thisGridCell = makeGridCell(wordAppCell.RowIndex, wordAppCell.ColumnIndex, wordAppCell);

                var crawlResult = new Base1List<GridCell>();

                int debugCounter = 0;

                //we have a cell to work with.
                while (thisGridCell != null)
                {
                    //get the snapshot
                    var snap = thisGridCell.Snapshot;

                    if (debugCounter++ > 1000 && Debugger.IsAttached)
                        Debugger.Break();
                    //have we seen this? then look to the next cell in the wordAppCells
                    if (visitedFastHashes.Contains(snap.FastHash))
                    {
                        if (Debugger.IsAttached) Debugger.Break();
                        break;
                    }

                    //this is where we add thic cell to the list for this set (row or col)
                    crawlResult.Add(thisGridCell);
                    //and add it to the stuff we've seen
                    visitedFastHashes.Add(thisGridCell.Snapshot.FastHash);

                    //now we probe in the direction to look for the next cell in this direction.
                    (int ProbeDepth, GridCell FoundCell)? probeResult = SendProbe(
                        thisGridCell,
                        crawlToVert,
                        crawlToHoriz,
                        visitedFastHashes);

                    //if we found something
                    if (probeResult != null)
                    {
                        //we can set this cell to the result of the probe.
                        //we also update 
                        thisGridCell = probeResult.Value.FoundCell;
                        //this is to update the coords, but I'm not sure we need it,
                        //since the GridCell and Word.Cell both have their own index references.

                        //if (crawlToVert != 0) wordTableRow += probeResult.Value.ProbeDepth;
                        //if (crawlToHoriz != 0) wordTableCol += probeResult.Value.ProbeDepth;
                    }
                    else
                    {
                        thisGridCell = null;
                    }
                }

                if (crawlResult.Count > 0)
                    outputLists.Add(crawlResult);
            }

            return outputLists;
        }

        private (int ProbeDepth, GridCell FoundCell)? SendProbe(
        GridCell thisGridCell,
        int probeToVert,
        int probeToHoriz,
        HashSet<ulong> visited)
        {
            var probeDepth = 1;

            while (true)
            {
                if (probeDepth > 500) // ← arbitrary sane limit
                {
                    if (Debugger.IsAttached) Debugger.Break();
                    return null; // 🛑 bail out
                }

                var res = thisGridCell.TryGetNeighbor(
                    probeToVert * probeDepth,
                    probeToHoriz * probeDepth,
                    out var probe);

                switch (res)
                {
                    case GridCell.GetNeighborResult.Success:
                        var probeHash = probe?.Snapshot.FastHash ?? 0;
                        if (!visited.Contains(probeHash))
                        {
                            return (probeDepth, probe);
                        }
                        break;

                    case GridCell.GetNeighborResult.COMExceptionThrownMergedCell:
                    //break;//broken. somethins this is actually out of range.
                    case GridCell.GetNeighborResult.OutOfBounds:
                    case GridCell.GetNeighborResult.SameCell:
                    case GridCell.GetNeighborResult.NotFound:
                        return null; // 🛑 bail out
                }

                probeDepth++;
            }
        }

    }





    public class GridCell
    {
        public Word.Cell COMCell { get; private set; }
        public bool IsMasterCell { get; private set; }
        public int GridRow { get; private set; }
        public int GridCol { get; private set; }
        public RangeSnapshot Snapshot { get; private set; }

        public GridCell(Word.Cell cell, int gridRow, int gridCol, bool isMasterCell)
        {
            COMCell = cell;
            GridRow = gridRow;
            GridCol = gridCol;
            IsMasterCell = isMasterCell;
            Snapshot = new RangeSnapshot(COMCell.Range);
        }

        public bool SameAs(GridCell other)
        {
            return other != null && Snapshot.FastMatch(other.Snapshot);
        }

        public enum GetNeighborResult
        {
            Success,
            OutOfBounds,
            SameCell,
            COMExceptionThrownMergedCell,
            COMExceptionThownOther,
            ExceptionThrown,
            NotFound
        }

        public GetNeighborResult TryGetNeighbor(int rowOffset, int colOffset, out GridCell neighbor)
        {
            neighbor = null;

            var table = COMCell.Range.Tables[1];
            var allCells = table.Range.Cells.ToList();

            _ = table.Cell(COMCell.RowIndex, COMCell.ColumnIndex);

            Word.Cell targetCell;
            try
            {
                targetCell = table.Cell(COMCell.RowIndex + rowOffset, COMCell.ColumnIndex + colOffset);
            }
            catch (COMException ex1)
            {
                if (ex1.Message.Contains("The requested member of the collection does not exist."))
                {

                    return GetNeighborResult.COMExceptionThrownMergedCell;
                }
                return GetNeighborResult.COMExceptionThownOther;
            }
            catch (Exception ex2)
            {
                return GetNeighborResult.ExceptionThrown;
            }

            var resolved = allCells.FirstOrDefault(c => RangeSnapshot.FastMatch(c.Range, targetCell.Range));
            if (resolved == null)
                return GetNeighborResult.NotFound;

            if (Snapshot.FastMatch(resolved.Range))
                return GetNeighborResult.SameCell;

            neighbor = new GridCell(resolved, GridRow + rowOffset, GridCol + colOffset, false);
            return GetNeighborResult.Success;

        }
    }
}
