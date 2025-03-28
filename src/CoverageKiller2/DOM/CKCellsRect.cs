using System;
using System.Collections.Generic;

namespace CoverageKiller2.DOM
{
    public class CKCellsRect : CKCells
    {
        public CKCellsRect(CKTable table, ICellRef<CKCellsRect> cellReference)
        //: base(table, cellReference)
        {
        }

        protected override IEnumerable<CKCell> BuildCells()
        {
            throw new NotImplementedException();
            //var cellsRect = Table.Converters.GetCells(Table, this, (ICellRef<CKCellsRect>)CellRef);
            //return cellsRect;
        }
    }

}

