using System;
using System.Collections.Generic;
using static CoverageKiller2.DOM.CKCellReference;

namespace CoverageKiller2.DOM
{
    public class CKRow : CKCells
    {
        internal int _lastIndex;

        public CKRow(IEnumerable<GridCell> rowCells) :
            base(null, null)
        {
            throw new NotImplementedException();
        }

        public CKRow(CKTable table, int index) :
            base(table, new CKCellReference(table, index, ReferenceTypes.Row))
        {
        }

        public CKRow(CKRows cKRows, int index) : base(null)
        {
            throw new NotImplementedException();
        }

        public CKCells Cells => this;

        // Property to access row's index
        public int Index => throw new NotImplementedException();


    }
}


