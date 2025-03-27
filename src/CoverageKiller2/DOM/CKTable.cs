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
    }

}
