using System;
using System.Collections;
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


    public class Base1JaggedList<T> : IReadOnlyList<Base1List<T>>
    {
        private readonly List<Base1List<T>> _rows = new List<Base1List<T>>();

        public Base1JaggedList()
        {
            _rows.Add(null); // index 0 = tombstone
        }

        public Base1JaggedList(List<List<T>> list)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));

            _rows = new List<Base1List<T>> { null }; // tombstone at index 0

            foreach (var row in list)
            {
                if (row == null)
                    throw new ArgumentException("Row list cannot contain null elements.", nameof(list));

                _rows.Add(new Base1List<T>(row));
            }
        }


        public void Add(Base1List<T> row)
        {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        public void Insert(int index, Base1List<T> row)
        {
            if (index < 1 || index > Count + 1)
                throw new ArgumentOutOfRangeException(nameof(index));
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Insert(index, row);
        }

        public void RemoveAt(int index)
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            _rows.RemoveAt(index);
        }

        public int IndexOf(Base1List<T> row)
        {
            int idx = _rows.IndexOf(row);
            return idx <= 0 ? -1 : idx;
        }

        public int Count => _rows.Count - 1;

        public Base1List<T> this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _rows[index];
            }
        }

        public IEnumerator<Base1List<T>> GetEnumerator()
        {
            for (int i = 1; i < _rows.Count; i++)
                yield return _rows[i];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }




    public class Base1List<T> : IReadOnlyList<T>
    {
        private readonly List<T> _items = new List<T>();

        public Base1List()
        {
            _items.Add(default); // index 0 = tombstone
        }

        public Base1List(Base1List<T> items)
        {
            _items.Add(default);
            _items.AddRange(items);
        }

        public Base1List(IEnumerable<T> items)
        {
            _items.Add(default);
            _items.AddRange(items);
        }
        public void Add(T item)
        {
            //if (item == null) throw new ArgumentNullException(nameof(item));
            _items.Add(item);
        }

        public void Insert(int index, T item)
        {
            if (index < 1 || index > Count + 1)
                throw new ArgumentOutOfRangeException(nameof(index));
            //if (item == null) throw new ArgumentNullException(nameof(item));
            _items.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            _items.RemoveAt(index);
        }

        public int IndexOf(T item)
        {
            int idx = _items.IndexOf(item);
            return idx <= 0 ? -1 : idx;
        }

        public int Count => _items.Count - 1;

        public T this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _items[index];
            }
        }
        public void PadToCount(int targetCount, T padWith = default)
        {

            if (targetCount < 1) throw new ArgumentOutOfRangeException(nameof(targetCount));

            int paddingNeeded = (targetCount + 1) - _items.Count;

            if (paddingNeeded > 0)
                _items.AddRange(Enumerable.Repeat(padWith, paddingNeeded));

        }
        public IEnumerator<T> GetEnumerator()
        {
            for (int i = 1; i < _items.Count; i++)
                yield return _items[i];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal void Clear()
        {
            _items.Clear();
            _items.Add(default);
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
