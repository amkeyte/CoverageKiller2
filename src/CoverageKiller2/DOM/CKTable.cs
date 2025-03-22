using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    public class CKTable : CKRange
    {
        // Instance Fields
        private CKTableGrid Grid { get; set; }

        // Constructors
        public CKTable(Word.Table table) : base(table.Range)
        {
            COMTable = table;
            Grid = CKTableGrid.GetInstance(table);

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

        // Methods
        public CKCell Cell(CKCellRefRect cellReference)
        {
            return new CKCell(Grid.GetCellAt(cellReference).COMCell);
        }



        /// <summary>
        /// Returns the collection of cells in the table that fall within the specified cell reference.
        /// </summary>
        /// <param name="cellReference">The cell reference specifying a range within the table.</param>
        /// <returns>A collection of CKCell objects.</returns>
        internal IEnumerable<CKCell> CellItems(CKCellRefRect cellReference)
        {
            throw new NotImplementedException();
        }
    }
}
