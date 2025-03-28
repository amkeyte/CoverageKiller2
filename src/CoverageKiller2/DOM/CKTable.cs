using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
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
            CKDocument doc = CKDocuments.GetByName(wordDoc.FullName);
            var foundTable = doc.Tables
                .Where(t => t.COMRange.Contains(wordRange))
                .FirstOrDefault();

            return foundTable ?? new CKTable(wordRange.Tables[1]);
        }

        // Instance Fields
        private CKTableGrid Grid { get; set; }
        private CKCellRefConverterService _converterService;

        // Constructors
        public CKTable(Word.Table table) : base(table.Range)
        {
            COMTable = table;
            Grid = CKTableGrid.GetInstance(table);
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

            internal CKTableGrid Grid => Table.Grid;

        }

        public CKCell Cell(CKCellRef cellRef)
        {
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            var wordCell = COMTable.Cell(cellRef.WordRow, cellRef.WordCol);
            return new CKCell(this, cellRef.Parent, wordCell, gridCellRef.X1, gridCellRef.Y1);
        }

        public IEnumerable<int> IndexesOf(Word.Cells wordCells)
        {
            var tableCellList = COMTable.Range.Cells.ToList();

            var matches = wordCells.AsEnumerable().
                Select(c => IndexOf(c, tableCellList));

            return matches;
        }

        internal int IndexOf(Word.Cell wordCell, List<Word.Cell> tableCells = null)
        {
            var tableCellList = tableCells ?? COMTable.Range.Cells.ToList(); //expensive!!
            return tableCellList.IndexOf(wordCell);
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
    }

}
