using System;
using System.Collections.Generic;

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



        public CKRow(CKRows cKRows, int index) : base(null)
        {
            throw new NotImplementedException();
        }

        public CKCells Cells => this;

        // Property to access row's index
        public int Index => throw new NotImplementedException();


    }
}


