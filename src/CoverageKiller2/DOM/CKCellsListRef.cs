using CoverageKiller2.DOM;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents an arbitrary collection of CKCell objects.
    /// </summary>
    public class CKCellsListRef : ICellRef<CKCells>
    {
        public CKTable Table { get; }
        public IEnumerable<int> CellIndexes { get; }

        public IDOMObject Parent => throw new NotImplementedException();

        public CKCellsListRef(IEnumerable<CKCell> cells)
        {
            if (cells == null)
                throw new ArgumentNullException(nameof(cells));
            Table = cells.FirstOrDefault()?.Table ??
                throw new NullReferenceException("null cell reference or table");
            if (cells.Any(c => c.Table != Table))
                throw new InvalidOperationException("Table Mismatch in cell references.");

            CellIndexes = cells.Select(c => c.CellRef.CellIndexes.First()).ToList();
        }

        public CKCellsListRef(Word.Cells wordCells)
        {
            if (wordCells == null) throw new ArgumentNullException(nameof(wordCells));
            if (wordCells.Count == 0) throw new ArgumentException("wordCells is empty");
            Table = CKTable.FromRange(wordCells[1].Range); //lazy constuctor for cached CKTables
            CellIndexes = Table.IndexesOf(wordCells); //Tables can cmpare to get the index of each cell

        }

        // Constructor for raw index list (e.g. FromWordCells)
        public CKCellsListRef(CKTable table, IEnumerable<int> cellIndexes)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            CellIndexes = cellIndexes?.ToList() ?? throw new ArgumentNullException(nameof(cellIndexes));
        }
    }

}

