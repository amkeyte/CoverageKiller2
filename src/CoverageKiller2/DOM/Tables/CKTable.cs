using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>


    public class CKTable : CKRange
    {
        public static CKTable FromRange(Word.Range wordRange)
        {
            if (wordRange == null) throw new ArgumentNullException(nameof(wordRange));
            if (wordRange.Tables.Count == 0)
                throw new ArgumentException($"{nameof(wordRange)} does not contain a table.");

            Word.Document wordDoc = wordRange.Document;
            CKDocument doc = CKDocuments.GetByCOMDocument(wordDoc);
            var foundTable = doc.Tables
                .Where(t => t.COMRange.Contains(wordRange))
                .FirstOrDefault();

            return foundTable ?? new CKTable(wordRange.Tables[1]);
        }

        // Instance Fields
        //private CKTableGrid Grid { get; set; }
        private CKTableGrid Grid => throw new NotImplementedException("DEBUG");//DEBUG
        private CKCellRefConverterService _converterService;

        // Constructors
        public CKTable(Word.Table table) : base(table.Range)
        {
            COMTable = table;

            //Grid = CKTableGrid.GetInstance(table);//DEBUG
            _converterService = new CKCellRefConverterService(this);
        }

        // Properties



        /// <summary>
        /// Remove from external references. Will be hidden.
        /// </summary>
        public Word.Table COMTable { get; private set; }

        /// <summary>
        /// Gets the rows of the table.
        /// </summary>
        public CKRows Rows => throw new NotImplementedException();



        /// <summary>
        /// Gets the columns of the table.
        /// </summary>
        public IEnumerable<CKColumn> Columns => throw new NotImplementedException();

        public CKCellRefConverterService Converters => _converterService;

        public class CKCellRefConverterService
        {
            public CKCellRefConverterService(CKTable table)
            {
                Table = table;
            }
            public CKTable Table { get; private set; }

            public CKTableGrid Grid => Table.Grid;

        }

        public CKCell Cell(CKCellRef cellRef)
        {
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            var wordCell = COMTable.Cell(cellRef.WordRow, cellRef.WordCol);
            return new CKCell(this, cellRef.Parent, wordCell, gridCellRef.Y1 + 1, gridCellRef.X1 + 1);
        }

        public IEnumerable<int> IndexesOf(Word.Cells wordCells)
        {
            var tableCellList = COMTable.Range.Cells.ToList();

            var matches = wordCells.AsEnumerable().
                Select(c => IndexOf(c, tableCellList));

            return matches;
        }

        public int IndexOf(Word.Cell wordCell, List<Word.Cell> tableCells = null)
        {
            var tableCellList = tableCells ?? COMTable.Range.Cells.ToList(); //expensive!!
            return tableCellList.FindIndex(c => c.Range.COMEquals(wordCell.Range)) + 1;
        }

        public IEnumerable<int> IndexesOf(CKCells cells)
        {
            IEnumerable<GridCell> masterCells = Grid.GetMasterCells();

            IEnumerable<GridCell> findCells = cells
                .Select(c => Converters.GetGridCellRef(c.CellRef))
                .Select(c => Grid.GetMasterCells(c).First());

            var matches = findCells.Select(c => IndexOf(c, masterCells));
            return matches;
        }

        private int IndexOf(GridCell c, IEnumerable<GridCell> masterCells = null)
        {
            var masterList = new List<GridCell>(masterCells ?? Grid.GetMasterCells());
            return masterList.IndexOf(c);
        }

        public CKCell Cell(int index)
        {
            var gridCellRef = Converters.GetGridCellRef(index);
            var wordCell = COMTable.Range.Cells[index];
            return new CKCell(this, this, wordCell, gridCellRef.Y1 + 1, gridCellRef.X1 + 1);
        }
    }
    /// <summary>
    /// Represents a collection of <see cref="CKTable"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    public class CKTables : ACKRangeCollection, IEnumerable<CKTable>
    {
        /// <summary>
        /// Gets the underlying Word.Tables COM object from the parent range.
        /// Note that there is only one Tables property, so calling back to it
        /// instead of storing a reference every time is acceptable.
        /// </summary>
        public Word.Tables COMTables => Parent.COMRange.Tables;

        /// <summary>
        /// Returns a string that represents the current <see cref="CKTables"/> instance.
        /// </summary>
        /// <returns>A string containing the count of tables.</returns>
        public override string ToString()
        {
            // Since CKRange doesn't provide a file path, we simply return the count.
            return $"CKTables [Count: {Count}]";
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTables"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        public CKTables(CKRange parent) : base(parent) { }

        /// <summary>
        /// Gets the number of tables in the associated range.
        /// </summary>
        public override int Count => COMTables.Count;

        /// <summary>
        /// optimize this if there's a big delay here.
        /// </summary>
        public override bool IsDirty => _isDirty || this.Any(x => x.IsDirty);

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the table to retrieve.</param>
        /// <returns>The <see cref="CKTable"/> at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the index is less than 1 or greater than the number of tables.
        /// </exception>
        public CKTable this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                }
                return new CKTable(COMTables[index]);
            }
        }

        /// <summary>
        /// Determines the one-based index of the specified <see cref="CKTable"/> in the collection.
        /// </summary>
        /// <param name="targetTable">The table to locate in the collection.</param>
        /// <returns>
        /// The one-based index of the table if found; otherwise, -1.
        /// </returns>
        //public int IndexOf(CKTable targetTable)
        //{
        //    for (int i = 1; i <= Count; i++)
        //    {
        //        var table = COMTables[i];

        //        // Compare by checking that both tables have the same start and end range
        //        if (table.Range.Start == targetTable.COMObject.Range.Start &&
        //            table.Range.End == targetTable.COMObject.Range.End)
        //        {
        //            return i;
        //        }
        //    }
        //    return -1;
        //}

        /// <summary>
        /// Returns an enumerator that iterates through the <see cref="CKTable"/> objects in the collection.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKTable"/> objects.</returns>
        public IEnumerator<CKTable> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator for the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
